# -*- coding: utf-8 -*-
# SHEIN AU 商品详情抓取模块（Chrome / Edge DevTools Protocol）
#
# 思路：
#   1) 用真实 Chrome + 独立用户目录。首次可能要过一次人机验证，之后 cookie 保留。
#   2) 数据直接取页面里的 window.gbRawData，字段齐全：
#        标题  modules.productInfo.goods_name
#        图集  modules.productInfo.allColorDetailImages[goods_id][].origin_image
#        变体  modules.saleAttr.multiLevelSaleAttribute.sku_list[].sku_sale_attr
#        价格  sku_list[].priceInfo.salePrice.amount
#        库存  sku_list[].stock
#   3) 图片 CDN 直接返回 JPEG 字节，存 .jpg 即可，不需要 Pillow。
#
# 依赖：websocket-client   （pip install websocket-client）

import json
import os
import re
import subprocess
import sys
import time
import urllib.parse
import urllib.request

try:
    import websocket
except Exception:
    websocket = None


IS_MAC = sys.platform == "darwin"
IS_WIN = sys.platform.startswith("win")

if IS_MAC:
    CHROME_CANDIDATES = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        os.path.expanduser("~/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"),
    ]
    EDGE_CANDIDATES = [
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        os.path.expanduser("~/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
    ]
elif IS_WIN:
    CHROME_CANDIDATES = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    EDGE_CANDIDATES = [
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
    ]
else:                                   # Linux 等
    CHROME_CANDIDATES = [
        "/usr/bin/google-chrome", "/usr/bin/google-chrome-stable",
        "/usr/bin/chromium", "/usr/bin/chromium-browser",
    ]
    EDGE_CANDIDATES = [
        "/usr/bin/microsoft-edge",
        "/usr/bin/microsoft-edge-stable",
    ]

DEFAULT_PORT = 9223


class Stopped(Exception):
    """用户点了停止"""


class NeedHuman(Exception):
    """希音要人机验证 / 要登录，脚本自己过不去"""


class NeedLanguage(NeedHuman):
    """页面不是英文，自动切换失败，需要人工处理语言。"""


def find_chrome():
    for p in CHROME_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return None


def find_edge():
    for p in EDGE_CANDIDATES:
        if p and os.path.exists(p):
            return p
    return None


def find_browser(name="Chrome"):
    """按用户选择查找 Chromium 浏览器；找不到时回退到另一个。"""
    if str(name or "").strip().lower() == "edge":
        return find_edge() or find_chrome()
    return find_chrome() or find_edge()


def is_connection_error(exc):
    """判断异常是否属于浏览器调试连接断开，而不是页面内容错误。

    Windows 上常见的是 ``[WinError 10054]``；websocket-client 自己的异常
    也会在这里统一识别。页面解析失败等普通错误不应被当成断线。
    """
    if exc is None:
        return False
    text = repr(exc).lower()
    markers = (
        "10053", "10054", "10060", "10061",
        "connection reset", "connection aborted", "connection refused",
        "remote host", "websocket", "socket", "broken pipe",
        "not connected", "connection is already closed",
        "connection closed", "connection lost", "timed out", "timeout",
    )
    if any(m in text for m in markers):
        return True
    return isinstance(exc, (ConnectionError, TimeoutError))


EXTRACT_JS = r"""
(function(){
  var out = { href: location.href, docTitle: document.title || '' };
  /* 评分 / 评价数：希音放在页面 meta 里（itemprop="ratingValue" / "reviewCount"） */
  try {
    var pick = function(names){
      for (var i = 0; i < names.length; i++) {
        var el = document.querySelector('[itemprop="' + names[i] + '"]');
        if (el) {
          var v = el.getAttribute('content') || el.innerText || '';
          v = (v || '').trim();
          if (v) return v;
        }
      }
      return '';
    };
    out.rating = pick(['ratingValue']);
    out.reviews = pick(['reviewCount', 'ratingCount']);
    out.ratingMax = pick(['bestRating']);
    /* 页面 JSON-LD 里每个变体的 sku（用来核对"我们读到的变体"是不是这一页的）*/
    out.pageSkus = [];
    try {
      var ld = document.querySelectorAll('script[type="application/ld+json"]');
      for (var i0 = 0; i0 < ld.length; i0++) {
        var txt0 = ld[i0].textContent || '';
        var re0 = /"sku"\s*:\s*"([^"]+)"/g, m0;
        while ((m0 = re0.exec(txt0)) !== null) {
          if (out.pageSkus.indexOf(m0[1]) < 0) out.pageSkus.push(m0[1]);
        }
      }
    } catch (e1) {}
    /* 兜底：JSON-LD 里的 aggregateRating */
    if (!out.rating) {
      var ss = document.querySelectorAll('script[type="application/ld+json"]');
      for (var i = 0; i < ss.length; i++) {
        var txt = ss[i].textContent || '';
        var m1 = txt.match(/"ratingValue"\s*:\s*"?([0-9.]+)"?/);
        if (m1) out.rating = m1[1];
        var m2 = txt.match(/"reviewCount"\s*:\s*"?([0-9]+)"?/);
        if (m2 && !out.reviews) out.reviews = m2[1];
        var m3 = txt.match(/"bestRating"\s*:\s*"?([0-9.]+)"?/);
        if (m3 && !out.ratingMax) out.ratingMax = m3[1];
        if (out.rating) break;
      }
    }
  } catch (e0) {}
  try {
    var d = window.gbRawData;
    if (d && d.modules) { out.ok = true; out.json = JSON.stringify(d); return JSON.stringify(out); }
  } catch (e) { out.err = String(e); }
  out.ok = false;
  try {
    var html = document.documentElement.outerHTML;
    var i = html.indexOf('window.gbRawData');
    if (i < 0) i = html.indexOf('"sku_list"');
    out.raw = (i >= 0) ? html.substring(Math.max(0, i - 2000), i + 400000) : '';
    var h1 = document.querySelector('h1');
    out.h1 = h1 ? (h1.innerText || '').replace(/\s+/g,' ').trim() : '';
    var og = document.querySelector('meta[property="og:title"]');
    out.ogTitle = og ? (og.getAttribute('content') || '') : '';
  } catch (e2) { out.err2 = String(e2); }
  return JSON.stringify(out);
})()
"""


# 分类/搜索列表页：把商品卡片扒出来
# 希音自己的报错页（Oops! Something went wrong / Page not found）——命中就自动刷新
OOPS_JS = r"""
(function(){
  try {
    var t = (document.title || '') + ' ';
    var b = document.body ? (document.body.innerText || '').slice(0, 400) : '';
    var s = (t + ' ' + b).toLowerCase();
    if (s.indexOf('oops') >= 0) return true;
    if (s.indexOf('something went wrong') >= 0) return true;
    if (s.indexOf('页面出错了') >= 0) return true;
    if (s.indexOf('page not found') >= 0) return true;
    return false;
  } catch (e) { return false; }
})()
"""


# 条款 / Cookie 弹窗（抓多了会跳出来）：自动点掉
DISMISS_JS = r"""
(function(){
  var body = (document.body ? (document.body.innerText || '') : '').toLowerCase();
  var looksLikeConsent = (body.indexOf('cookie') >= 0) ||
                         (body.indexOf('terms & conditions') >= 0) ||
                         (body.indexOf('terms and conditions') >= 0) ||
                         (body.indexOf('privacy') >= 0);
  if (!looksLikeConsent) return '';
  /* 只认「同意」类按钮；CONTINUE / OK 不点（登录页上那个 CONTINUE 不能乱点）*/
  var words = ['accept all', 'accept', 'agree', 'i agree', 'got it', 'allow all'];
  var els = document.querySelectorAll('button, a, div[role="button"], span[role="button"]');
  for (var i = 0; i < els.length; i++) {
    var el = els[i];
    if (el.offsetParent === null) continue;          // 不可见的跳过
    var t = (el.innerText || '').replace(/\s+/g, ' ').trim().toLowerCase();
    if (!t || t.length > 24) continue;
    for (var j = 0; j < words.length; j++) {
      if (t === words[j]) {
        try { el.click(); } catch (e) {}
        return t;
      }
    }
  }
  return '';
})()
"""


LANGUAGE_STATE_JS = r"""
(function(){
  try {
    var lang = (document.documentElement.lang || '').toLowerCase();
    var body = document.body ? (document.body.innerText || '').slice(0, 20000) : '';
    var zh = lang.indexOf('zh') === 0 ||
      /加入购物车|购物车|尺码|客户评价|关于品牌|商品详情/.test(body);
    return JSON.stringify({zh: zh, lang: lang, href: location.href});
  } catch (e) {
    return JSON.stringify({zh: false, lang: '', href: ''});
  }
})()
"""


CLICK_ENGLISH_JS = r"""
(function(){
  function visible(el){
    try {
      var r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0;
    } catch (e) { return false; }
  }
  var els = document.querySelectorAll('a,button,[role="button"],div,span,li');
  for (var i = 0; i < els.length; i++) {
    var el = els[i];
    if (!visible(el)) continue;
    var t = (el.innerText || '').replace(/\s+/g, ' ').trim();
    if (t === 'English' || t === '英语') {
      try { el.click(); return 'english'; } catch (e) {}
    }
  }
  // 没有看到 English 选项时，先打开页面右上角的语言菜单。
  for (var j = 0; j < els.length; j++) {
    var x = els[j];
    if (!visible(x)) continue;
    var label = ((x.getAttribute('aria-label') || '') + ' ' +
                 (x.getAttribute('title') || '') + ' ' +
                 (x.innerText || '')).toLowerCase();
    if (/language|语言|語言/.test(label)) {
      try { x.click(); return 'opened'; } catch (e) {}
    }
  }
  return '';
})()
"""


CATEGORY_JS = r"""
(function(){
  var out = { href: location.href, cards: 0, items: [] };
  var cards = document.querySelectorAll('div[role="listitem"]');
  out.cards = cards.length;
  for (var i = 0; i < cards.length; i++) {
    var c = cards[i];
    var a = c.querySelector('a[href*="-p-"]');
    if (!a) continue;
    var h = a.getAttribute('href') || '';
    var m = h.match(/-p-(\d+)\.html/);
    if (!m) continue;
    var nameEl = c.querySelector('.bs-product-card__goods-name-text');
    var priceEl = c.querySelector('.bs-product-card__price-cart-wrapper');
    var ptext = priceEl ? (priceEl.innerText || '') : '';
    /* 价格：澳洲站是 AU$12.95，美国站是 $8.53，这里都不挑币种 */
    var pm = ptext.match(/(AU|US|NZ|CA|SG|UK)?\s*\$\s*([0-9]+(?:\.[0-9]+)?)/);
    var imgEl = c.querySelector('.crop-image-container[data-before-crop-src]');
    var img = imgEl ? (imgEl.getAttribute('data-before-crop-src') || '') : '';
    if (img.indexOf('//') === 0) img = 'https:' + img;
    /* 卡片上的销量 / 爆款标签（都在卡片文本里）*/
    var cardText = (c.innerText || '').replace(/\s+/g, ' ');
    var sm = cardText.match(/([\d.,]+\s*[kKmM]?\+?)\s*sold/i);
    var lm = cardText.match(/#\s*\d+\s*Bestseller[^#]{0,40}/i);
    if (!lm) lm = cardText.match(/\b(HOT|Trending|Best Seller)\b/i);
    out.items.push({
      goods_id: m[1],
      url: (h.indexOf('http') === 0 ? h : location.origin + h).split('?')[0],
      title: nameEl ? (nameEl.innerText || '').trim()
                    : (a.getAttribute('aria-label') || ''),
      price: pm ? parseFloat(pm[2]) : null,
      currency: pm ? ((pm[1] || '') + '$') : '',
      sold: sm ? sm[1].trim() : '',
      label: lm ? lm[0].trim() : '',
      img: img
    });
  }
  return JSON.stringify(out);
})()
"""


def balanced(text, start_idx):
    open_ch = text[start_idx]
    close_ch = "]" if open_ch == "[" else "}"
    depth = 0
    in_str = False
    esc = False
    for i in range(start_idx, len(text)):
        ch = text[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return text[start_idx:i + 1]
    return None


def parse_raw_blob(raw):
    if not raw:
        return None
    i = raw.find("window.gbRawData")
    if i >= 0:
        j = raw.find("{", i)
        if j >= 0:
            blob = balanced(raw, j)
            if blob:
                try:
                    return json.loads(blob)
                except Exception:
                    pass
    return None


def with_url_params(url, params=None, drop=()):
    """在链接上增删查询参数（价格筛选 min_price/max_price、翻页 page 都用这个）"""
    p = urllib.parse.urlsplit(url)
    q = [(k, v) for k, v in urllib.parse.parse_qsl(p.query, keep_blank_values=True)
         if k.lower() not in {str(d).lower() for d in drop}]
    for k, v in (params or {}).items():
        q.append((k, str(v)))
    return urllib.parse.urlunsplit(
        (p.scheme, p.netloc, p.path, urllib.parse.urlencode(q), p.fragment))


def _num(v, default=0.0):
    try:
        return float(v)
    except Exception:
        return default


def parse_product(data):
    """把 gbRawData 解析成统一结构"""
    out = {"goods_id": "", "skc_name": "", "title": "", "images": [],
           "variants": [], "price_min": None, "price_max": None, "error": "",
           "store": "", "colors": [], "color": "", "currency": "",
           "rating": None, "reviews": 0, "is_local": False, "quick_ship": False,
           "page_skus": [], "url_goods_id": ""}
    mods = (data or {}).get("modules") or {}
    pi = mods.get("productInfo") or {}
    out["goods_id"] = str(pi.get("goods_id") or "")
    out["skc_name"] = str(pi.get("goods_sn") or "")
    out["title"] = (pi.get("goods_name") or "").strip()
    out["store"] = str((mods.get("storeInfo") or {}).get("title") or "").strip()
    # Local（本地仓）/ 快速发货：productInfo 有，SKU 级也有
    local_flags = [str(pi.get("isLocalProduct") or "0")]
    quick_flags = []
    for _s in ((mods.get("saleAttr") or {}).get("multiLevelSaleAttribute") or {}).get("sku_list") or []:
        if isinstance(_s, dict):
            local_flags.append(str(_s.get("isLocalProduct") or "0"))
            quick_flags.append(str(_s.get("isSupportQuickShip") or "0"))
    out["is_local"] = ("1" in local_flags)
    out["quick_ship"] = ("1" in quick_flags or "true" in [f.lower() for f in quick_flags])

    allc = pi.get("allColorDetailImages") or {}
    ci = pi.get("currentSkcImgInfo") or {}
    seq = None
    if isinstance(allc, dict):
        seq = allc.get(out["goods_id"])
        if seq is None and allc:
            seq = list(allc.values())[0]
    first = []
    for it in (seq or []):
        if isinstance(it, dict):
            u = it.get("origin_image") or it.get("image") or ""
        else:
            u = it if isinstance(it, str) else ""
        if u:
            first.append(u)
    tail = list(ci.get("skcImages") or []) + list(ci.get("spuImages") or [])
    if pi.get("goods_img"):
        tail.append(pi["goods_img"])
    imgs = []
    for u in first + tail:
        if isinstance(u, str) and u and u not in imgs:
            imgs.append(u)
    out["images"] = [u if u.startswith("http") else "https:" + u for u in imgs]

    sa = mods.get("saleAttr") or {}
    ml = sa.get("multiLevelSaleAttribute") or {}
    if not out["skc_name"]:
        out["skc_name"] = str(ml.get("skc_name") or "")
    variants = []
    for s in (ml.get("sku_list") or []):
        if not isinstance(s, dict):
            continue
        price = ((s.get("priceInfo") or {}).get("salePrice") or {}).get("amount")
        if price is None:
            price = ((s.get("price") or {}).get("salePrice") or {}).get("amount")
        attrs = s.get("sku_sale_attr") or []
        names = [a.get("attr_value_name") for a in attrs if a.get("attr_value_name")]
        variants.append({
            "sku_code": str(s.get("sku_code") or ""),
            "name": ("-".join(names).strip("-") or "One Size"),
            "price": _num(price, None),
            "stock": int(_num(s.get("stock"), 0)),
        })
    if not variants:
        price = ((mods.get("priceInfo") or {}).get("salePrice") or {}).get("amount")
        if price is not None:
            variants.append({"sku_code": out["skc_name"], "name": "One Size",
                             "price": _num(price, None),
                             "stock": int(_num(pi.get("stock"), 0))})
    out["variants"] = variants
    prices = [v["price"] for v in variants if v["price"]]
    if prices:
        out["price_min"] = min(prices)
        out["price_max"] = max(prices)

    # ---- 颜色（希音这里常常是「一个颜色 = 一个独立商品」，各有自己的 goods_id）----
    main = sa.get("mainSaleAttribute") or {}
    colors = []
    for c in (main.get("info") or []):
        if not isinstance(c, dict):
            continue
        colors.append({
            "id": str(c.get("attr_value_id") or ""),
            "name": (c.get("attr_value") or "").strip(),
            "goods_id": str(c.get("goods_id") or ""),
            "goods_sn": str(c.get("goods_sn") or ""),
            "url_name": (c.get("goods_url_name") or "").strip(),
        })
    # 币种（AU$ / $ / € …），方便判断抓的是哪个站
    sym = ""
    _pi = mods.get("priceInfo") or {}
    _sp = ((_pi.get("salePrice") or {}).get("amountWithSymbol")
           or ((ml.get("sku_list") or [{}])[0].get("priceInfo", {})
               .get("salePrice", {}) or {}).get("amountWithSymbol") or "")
    m = re.match(r"^(\D+?)\s*[\d.,]", str(_sp))
    if m:
        sym = m.group(1).strip()
    out["currency"] = sym
    out["colors"] = colors
    cur = ""
    for c in colors:
        if c["goods_id"] and c["goods_id"] == out["goods_id"]:
            cur = c["name"]
            break
    if not cur and colors:
        cur = colors[0]["name"]
    out["color"] = cur
    return out


class SheinScraper:
    """浏览器常驻 + 逐个商品抓取"""

    def __init__(self, profile_dir, port=DEFAULT_PORT, chrome=None, log=print,
                 browser_name="Chrome", extra_args=None):
        self.profile = profile_dir
        self.port = port
        self.browser_name = str(browser_name or "Chrome")
        self.chrome = chrome or find_browser(self.browser_name)
        self.extra_args = list(extra_args or [])
        self.log = log
        self.ws = None
        self._id = 0
        self._proc = None
        self._last_connect_error = ""

    def port_alive(self, timeout=1.0):
        """这个端口上有没有一个能用的 DevTools（不是"端口被占用"就算）"""
        import socket
        try:
            s = socket.create_connection(("127.0.0.1", self.port), timeout=timeout)
            s.close()
        except Exception:
            return False
        try:
            with urllib.request.urlopen(
                    "http://127.0.0.1:%d/json/version" % self.port, timeout=2) as r:
                return b"Browser" in r.read(200)
        except Exception:
            return False

    @staticmethod
    def port_free(port):
        """端口能不能被我们绑（被别的进程占着就返回 False）"""
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind(("127.0.0.1", port))
            return True
        except Exception:
            return False
        finally:
            s.close()

    def start(self, visible=True, force=False):
        # 1) 先看配置端口上有没有能用的 DevTools（有就直接接管）
        # 2) 没有就找一个真正空闲的端口（可能被别的 Chrome 占着但没服务）
        use_port = None
        for cand in range(self.port, self.port + 6):
            self.port = cand
            if not force and self.port_alive():
                self.log("    检测到调试 %s（端口 %d），直接接管"
                         % (self.browser_name, cand))
                return True
            if self.port_free(cand):
                use_port = cand
                break
            self.log("    端口 %d 被别的程序占着（不是可用的 DevTools），换下一个" % cand)
        if use_port is None:
            self.log("    ！%d~%d 都被占用了，请关掉一些浏览器或改界面的「调试端口」"
                     % (self.port - 5, self.port))
            return False
        self.port = use_port
        # 同一个档案已经有残留的浏览器在跑的话，新实例只会往旧进程里开标签，
        # 调试端口永远起不来（表现就是窗口一直冒出来但脚本连不上）——先清掉它们
        stuck = self._profile_chrome_pids()
        if stuck:
            self.log("    有 %d 个 %s（同一个档案）占着，先关掉 …"
                     % (len(stuck), self.browser_name))
            self._kill_pids(stuck)
            # 只等真正退出，不再固定睡 3 秒；最多等 6 秒。
            self._wait_profile_exit(6.0)
        if not self.chrome:
            self.log("    找不到 %s，无法启用抓取" % self.browser_name)
            return False
        os.makedirs(self.profile, exist_ok=True)
        # 清掉异常退出留下的单例锁。只在确认没有同档案进程后清理。
        self._clear_profile_singletons()
        # 同一个档案的残留浏览器会"抢走"新启动的请求（新窗口进旧进程、没有调试端口）
        for attempt in range(2):
            args = [self.chrome,
                    "--remote-debugging-port=%d" % self.port,
                    "--user-data-dir=%s" % self.profile,
                    *self.extra_args,
                    "--no-first-run", "--no-default-browser-check",
                    "--disable-popup-blocking",
                    "--window-size=1360,900", "--window-position=120,60",
                    "--remote-allow-origins=*", "about:blank"]
            if not visible:
                args.insert(1, "--headless=new")
            try:
                self._proc = subprocess.Popen(
                    args,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            except Exception as exc:
                self.log("    ！启动 %s 失败: %s" % (self.browser_name, exc))
                return False
            # 首次启动最多等 20 秒；失败就立即返回，不再做多轮长时间重试。
            for _ in range(40):
                if self.port_alive():
                    self.log("    %s 已在端口 %d 就绪"
                             % (self.browser_name, self.port))
                    return True
                if self._proc is not None and self._proc.poll() is not None:
                    self.log("    ！%s 启动后马上退出（退出码 %s）"
                             % (self.browser_name, self._proc.returncode))
                    return False
                time.sleep(0.5)
            if attempt == 0:
                # 只做一次短重试：清掉这次启动留下的进程，换一个端口再试。
                stuck = self._profile_chrome_pids()
                if stuck:
                    self.log("    端口没起来，先清掉 %d 个同档案的残留 %s …"
                             % (len(stuck), self.browser_name))
                    self._kill_pids(stuck)
                    self._wait_profile_exit(5.0)
                for cand in range(self.port, self.port + 6):
                    if self.port_free(cand):
                        self.port = cand
                        break
        self.log("    ！%s 起来了但调试端口没通（端口 %d）。"
                 "多半是同一个档案还开着别的窗口，把它关掉再跑；"
                 "或者点界面上的「换新会话」" % (self.browser_name, self.port))
        return False

    def _wait_profile_exit(self, timeout=6.0):
        """等同一档案的浏览器真正退出，返回是否已清空。"""
        end = time.time() + max(0.0, float(timeout))
        while time.time() < end:
            if not self._profile_chrome_pids():
                return True
            time.sleep(0.5)
        return not bool(self._profile_chrome_pids())

    def _clear_profile_singletons(self):
        """清理档案根目录下的 Chromium 单例文件。

        Windows 下这几个文件有时会在浏览器异常退出后残留，导致新进程
        直接退出或调试端口永远起不来。只动我们自己的 profile 根目录。
        """
        if self._profile_chrome_pids():
            return
        for name in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            p = os.path.join(self.profile, name)
            try:
                if os.path.lexists(p):
                    os.remove(p)
            except Exception:
                pass

    def _profile_chrome_pids(self):
        """用我们档案启动的浏览器进程号（只认这个档案）
        Windows 用 CIM 查命令行；macOS / Linux 用 pgrep -f"""
        try:
            if IS_WIN:
                exe_name = os.path.basename(self.chrome or "").lower()
                is_edge = "edge" in self.browser_name.lower() or exe_name == "msedge.exe"
                if exe_name not in ("chrome.exe", "msedge.exe"):
                    exe_name = "msedge.exe" if is_edge else "chrome.exe"
                if is_edge:
                    # 正常 Edge 主进程通常只显示 --profile-directory=Default，
                    # 不带完整的 User Data 路径；必须同时匹配这两个特征。
                    ps = (
                        "Get-CimInstance Win32_Process -Filter \"Name='%s'\" | "
                        "Where-Object { $_.CommandLine -like '*%s*' -or "
                        "$_.CommandLine -like '*--profile-directory=Default*' } | "
                        "Select-Object -ExpandProperty ProcessId"
                        % (exe_name, self.profile)
                    )
                else:
                    ps = ("Get-CimInstance Win32_Process -Filter \"Name='%s'\" | "
                          "Where-Object { $_.CommandLine -like '*%s*' } | "
                          "Select-Object -ExpandProperty ProcessId"
                          % (exe_name, self.profile))
                r = subprocess.run(["powershell", "-NoProfile", "-Command", ps],
                                   capture_output=True, text=True, timeout=30)
                return [int(x) for x in (r.stdout or "").split() if x.strip().isdigit()]
            r = subprocess.run(["pgrep", "-f", self.profile],
                               capture_output=True, text=True, timeout=30)
            return [int(x) for x in (r.stdout or "").split() if x.strip().isdigit()]
        except Exception:
            return []

    def _kill_pids(self, pids):
        for pid in pids:
            try:
                if IS_WIN:
                    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"],
                                   capture_output=True, timeout=15)
                else:
                    os.kill(pid, 9)
            except Exception:
                pass

    def connect(self):
        if websocket is None:
            raise RuntimeError("缺少 websocket-client，请执行:  pip install websocket-client")
        with urllib.request.urlopen(
                "http://127.0.0.1:%d/json" % self.port, timeout=10) as r:
            targets = json.loads(r.read().decode("utf-8", "ignore"))
        pages = [t for t in targets if t.get("type") == "page"]
        if not pages:
            raise RuntimeError("%s 里没有可用标签页" % self.browser_name)
        self.ws = websocket.create_connection(
            pages[0]["webSocketDebuggerUrl"], timeout=180,
            max_size=64 * 1024 * 1024)
        self._id = 0
        self.cmd("Page.enable")
        self.cmd("Runtime.enable")
        return True

    def close(self):
        try:
            if self.ws:
                self.ws.close()
        except Exception:
            pass
        finally:
            self.ws = None

    def reconnect(self, force=False):
        """浏览器调试连接断开后，重新接管或重开专用实例。

        先在原端口尝试重新连接；如果端口已经没有 DevTools 服务，或者
        指定 force，就清掉同档案的旧浏览器进程并重新启动一个。
        """
        self.close()
        self._proc = None
        if not force and self.port_alive():
            try:
                self.connect()
                self.log("    %s 调试连接已恢复（端口 %d）"
                         % (self.browser_name, self.port))
                return True
            except Exception as exc:
                self._last_connect_error = str(exc)
                self.log("    原端口重连失败: %s" % exc)
        if not self.start(force=True):
            self.log("    ！%s 重开失败，停止本批，避免继续写失败记录"
                     % self.browser_name)
            return False
        try:
            self.connect()
            self.log("    %s 已重开并重新连接（端口 %d）"
                     % (self.browser_name, self.port))
            return True
        except Exception as exc:
            self._last_connect_error = str(exc)
            self.log("    ！%s 重连失败: %s" % (self.browser_name, exc))
            return False

    def cmd(self, method, params=None, timeout=60):
        if self.ws is None:
            raise ConnectionError("%s 调试连接不存在" % self.browser_name)
        try:
            self.ws.settimeout(timeout)
        except Exception:
            pass
        self._id += 1
        mid = self._id
        self.ws.send(json.dumps({"id": mid, "method": method,
                                 "params": params or {}}))
        deadline = time.time() + timeout
        while True:
            if time.time() > deadline:
                raise RuntimeError("CDP 超时: %s" % method)
            msg = json.loads(self.ws.recv())
            if msg.get("id") == mid:
                return msg

    def eval(self, expr):
        r = self.cmd("Runtime.evaluate",
                     {"expression": expr, "returnByValue": True,
                      "awaitPromise": True})
        return r.get("result", {}).get("result", {}).get("value")

    def safe_eval(self, expr):
        try:
            return self.eval(expr)
        except Exception as exc:
            # 不要吞掉断线：上层需要立即停止/重连，而不是继续访问死连接。
            if is_connection_error(exc):
                raise
            return None

    def dismiss_overlays(self, state=None):
        """遇到条款 / Cookie 同意弹窗就自动点掉（点不到会让上层提示人工处理）"""
        state = state if state is not None else {}
        now = time.time()
        if now - state.get("last", 0) < 4:
            return ""
        state["last"] = now
        href = (self.safe_eval("location.href") or "").lower()
        if ("/user/auth" in href) or ("login" in href) or ("risk/challenge" in href):
            return ""            # 登录页 / 验证页 不要乱点
        hit = self.safe_eval(DISMISS_JS) or ""
        if hit:
            self.log("    自动点掉了弹窗：%s" % hit.upper())
        return hit

    def ensure_english(self, stop=None):
        """确保当前页面是英文；能自动切换就切换，失败则要求人工处理。

        SHEIN 会按地区/访客 Cookie 自动切到中文。先尝试点击英文选项，
        再尝试点开语言菜单后选择 English；连续失败就抛 NeedLanguage，
        不允许继续用中文页面解析商品。
        """
        for attempt in range(3):
            if stop is not None and stop.is_set():
                raise Stopped()
            try:
                state = json.loads(self.safe_eval(LANGUAGE_STATE_JS) or "{}")
            except Exception:
                state = {}
            if not state.get("zh"):
                return True
            if attempt == 0:
                self.log("    检测到中文页面，正在尝试切换 English …")
            clicked = self.safe_eval(CLICK_ENGLISH_JS) or ""
            if clicked:
                self.log("    语言操作：%s，等待页面刷新 …" % clicked)
            time.sleep(1.5)
        raise NeedLanguage(
            "页面仍是中文，自动切换 English 失败；请在浏览器右上角 "
            "Language 选择 English 后继续")

    def fetch_product(self, url, captcha_wait=300, ready_wait=35, log_every=15,
                      stop=None, wall_action="wait"):
        self.goto(url, "!!(window.gbRawData && window.gbRawData.modules)",
                  captcha_wait, ready_wait, log_every,
                  "等待页面数据超时，按当前状态解析", stop=stop,
                  wall_action=wall_action)
        self.ensure_english(stop=stop)
        time.sleep(1.0)
        raw = self.safe_eval(EXTRACT_JS)
        pack = json.loads(raw) if raw else {}
        # 星级是页面后加载的：头几次没读到就等一会儿再读一遍
        for _ in range(4):
            if pack.get("rating") or not pack.get("ok"):
                break
            time.sleep(2.0)
            raw2 = self.safe_eval(EXTRACT_JS)
            if raw2:
                pack = json.loads(raw2)
        if not raw:
            if not pack:
                raise RuntimeError("页面数据提取失败")
        data = None
        if pack.get("ok") and pack.get("json"):
            try:
                data = json.loads(pack["json"])
            except Exception:
                data = None
        if data is None:
            data = parse_raw_blob(pack.get("raw") or "")
        info = parse_product(data)
        if not info["title"] and pack.get("h1"):
            info["title"] = pack["h1"]
        if not info["title"] and pack.get("ogTitle"):
            info["title"] = pack["ogTitle"]
        # 评分 / 评价数（页面 meta，不在 gbRawData 里）
        info["rating_max"] = _num(pack.get("ratingMax"), 5.0) or 5.0
        r = _num(pack.get("rating"), None)
        info["rating"] = r
        info["reviews"] = int(_num(pack.get("reviews"), 0))
        info["page_skus"] = pack.get("pageSkus") or []
        m = re.search(r"-p-(\d+)\.html", url or "")
        info["url_goods_id"] = m.group(1) if m else ""
        info["href"] = pack.get("href") or url
        return info

    def goto(self, url, ready_expr, captcha_wait=300, ready_wait=35, log_every=15,
             tip="页面就绪等待超时", stop=None, wall_action="wait"):
        """打开页面并等它就绪。
        连不上（ERR_CONNECTION_CLOSED / 被限流）会自动等一会儿重试，最多 3 次。
        wall_action：遇到验证/登录墙怎么办——
          wait = 等你手动处理；skip = 立刻报错跳过；cool = 不等，交给上层歇会儿重试"""
        for attempt in range(3):
            self.cmd("Page.navigate", {"url": url})
            t0 = time.time()
            last_tip = 0.0
            last_alive = 0.0
            broken = False
            oops = 0
            overlay_state = {}
            while True:
                time.sleep(1.0)
                if stop is not None and stop.is_set():
                    raise Stopped()
                self.dismiss_overlays(state=overlay_state)      # 条款 / Cookie 弹窗自动点掉
                # 希音自己的报错页（Oops / Something went wrong）：自动刷新重试
                if self.safe_eval(OOPS_JS):
                    oops += 1
                    if oops > 3:
                        raise RuntimeError("页面一直是希音的报错页（Oops），刷新 3 次也没好")
                    self.log("      页面是希音的报错页（Oops），自动刷新重试 %d/3 …" % oops)
                    self.cmd("Page.reload", {"ignoreCache": True})
                    time.sleep(4)
                    continue
                href = (self.safe_eval("location.href") or "").lower()
                if href.startswith("chrome-error://"):
                    broken = True
                    break
                # 人机验证 / 登录墙：等你手动处理，脚本一直等
                need_human, why = False, ""
                if ("risk/challenge" in href) or ("captcha" in href):
                    need_human, why = (
                        True, "希音弹人机验证，请在浏览器窗口里过一下")
                elif ("/user/auth" in href) or ("/login" in href):
                    need_human, why = True, ("希音要求登录，请在浏览器窗口里登录一次"
                                             "（登录状态会记住）")
                if need_human:
                    if wall_action == "skip":
                        raise NeedHuman(why)
                    if time.time() - t0 > captcha_wait:
                        raise NeedHuman("%s —— 等太久（%d 秒）" % (why, captcha_wait))
                    if time.time() - last_tip > log_every:
                        self.log("    [需要人工] %s，脚本在等 …" % why)
                        last_tip = time.time()
                    continue
                if self.safe_eval(ready_expr):
                    return True
                if time.time() - last_alive > 6:
                    self.log("      页面加载中 …（已 %d 秒）" % int(time.time() - t0))
                    last_alive = time.time()
                if time.time() - t0 > ready_wait:
                    self.log("    警告：%s" % tip)
                    return False
            if broken and attempt < 2:
                wait = 10 * (attempt + 1)
                self.log("    网站连不上（连接被断开），%d 秒后重试第 %d 次 …"
                         % (wait, attempt + 2))
                for _ in range(wait):
                    if stop is not None and stop.is_set():
                        raise Stopped()
                    time.sleep(1)
                continue
            if broken:
                raise RuntimeError("连不上 au.shein.com（连接被断开）——"
                                   "多半是被临时限流了，歇一会儿或换个网络再跑")
        return False

    def fetch_category(self, url, scroll_rounds=8, stop_after=None,
                       captcha_wait=300, log=None, stop=None, enough=None,
                       wall_action="wait", max_pages=1):
        """从分类/搜索列表页扒商品卡片，返回 [{goods_id,url,title,price,img}]
        max_pages > 1 时会翻页（给链接加 page=2、3…），每页先滚到底把懒加载滚出来"""
        log = log or self.log
        items = []
        seen = set()
        add = 0
        page = 1
        while page <= max(1, int(max_pages or 1)):
            target = url if page == 1 else with_url_params(url, {"page": page})
            if page > 1:
                log("      翻到第 %d 页（累计候选 %d 个）…" % (page, len(items)))
            self.goto(target,
                      "(document.querySelectorAll('div[role=\"listitem\"]').length > 3)",
                      captcha_wait, 45, 15, "分类页加载超时，按当前内容解析",
                      stop=stop, wall_action=wall_action)
            self.ensure_english(stop=stop)
            for retry in range(2):
                if self.safe_eval("document.querySelectorAll('div[role=\"listitem\"]')"
                                  ".length") > 0:
                    break
                log("      分类页没出商品，重新加载一次 …")
                self.cmd("Page.reload", {"ignoreCache": True})
                time.sleep(6)

            add = 0
            last_count = -1
            stagnant = 0
            for r in range(scroll_rounds):
                if stop is not None and stop.is_set():
                    raise Stopped()
                raw = self.safe_eval(CATEGORY_JS)
                if raw:
                    try:
                        pack = json.loads(raw)
                    except Exception:
                        pack = {}
                    for it in (pack.get("items") or []):
                        gid = it.get("goods_id")
                        if not gid or gid in seen:
                            continue
                        seen.add(gid)
                        items.append(it)
                        add += 1
                log("      已见商品 %d 个（候选池，够要抓的数量就停）" % len(items))
                if enough is not None:
                    try:
                        ok = enough(items)
                    except Exception:
                        ok = False
                    if ok:
                        log("      候选够了，不再翻页")
                        return items
                if stop_after and len(items) >= stop_after:
                    return items
                if len(items) <= last_count:
                    stagnant += 1
                    if stagnant >= 3:
                        log("      本页不再加载更多商品")
                        break
                else:
                    stagnant = 0
                last_count = len(items)
                self.safe_eval("window.scrollTo(0, document.body.scrollHeight)")
                self.safe_eval("var e=document.scrollingElement||document.documentElement;"
                               "e.scrollTop=e.scrollHeight;")
                self.dismiss_overlays()
                time.sleep(3.5)

            if add == 0:
                log("      这一页没出新的商品，说明没有下一页了")
                break
            page += 1
        return items
