"""Build data/earnings_calendar.csv: past and upcoming earnings for every Equity Stock in universe.csv.

One row per stock x fiscal quarter, from the start of fiscal year FIRST_FY:
  event_date      results date: Wind stm_issuingdate when reported, else Wind's expected date
                  (stm_predict_issuingdate); for the next report, the company's announced date when
                  Nasdaq.com (Zacks) has one, and Nasdaq's or Yahoo's estimate when Wind has no date
  date_status     Confirmed (reported, or the company has announced the date) or Estimate
  date_status_source  what the status rests on: Wind (reported) or Nasdaq (Zacks); blank for later quarters
  date_yahoo      Yahoo's date for the next report, when it differs from the date used
  session         from universe.csv earnings_session: AMC after the US close, BMO before the open
  currency        the company's reporting currency (Yahoo financialCurrency): TWD for TSMC, KRW for SK hynix,
                  EUR for Nokia, else USD. Revenue figures are in it; EPS figures are per US-listed share
                  (per ADR) in USD
  rev_actual_bn   quarterly revenue, bn (Wind oper_rev, year-to-date figures differenced)
  eps_gaap        quarterly diluted GAAP EPS (Wind eps_diluted, year-to-date figures differenced); blank for
                  non-USD reporters, where Wind's figure is per ordinary share in the home currency
  eps_adj_actual  adjusted EPS as Yahoo reports it, the basis of the EPS consensus (last 4 quarters)
  eps_cons_*      quarterly EPS consensus: mean, low, high, analysts (Yahoo; next quarter and last 4)
  rev_cons_*_bn   quarterly revenue consensus, bn (Yahoo; next quarter only)
  fy_*            consensus for the fiscal year the quarter falls in: adjusted EPS and revenue (bn),
                  mean, low, high, analysts (Yahoo); upcoming quarters only
  category        calendar category: Equity for these rows (economic releases will be Economic)

All consensus figures come from Yahoo, on one basis (adjusted EPS); dates and actuals from Wind,
whether an upcoming date is confirmed from Nasdaq.com (Zacks).
Run alone or from build_data.py; needs Wind Terminal open and logged in.
"""
import csv
import http.cookiejar
import json
import re
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "data" / "earnings_calendar.csv"
SGT = timezone(timedelta(hours=8))
FIRST_FY = 2026
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/130 Safari/537.36"
MONTHS = {m: i for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], start=1)}
COLS = ["event_id", "code", "ticker", "name", "category", "kind", "event_date", "session", "status", "date_status",
        "date_status_source", "date_source",
        "date_yahoo", "fiscal_period", "period_end", "currency", "rev_actual_bn", "eps_gaap", "eps_adj_actual",
        "eps_surprise_pct", "rev_cons_bn", "rev_cons_low_bn", "rev_cons_high_bn", "rev_cons_n", "eps_cons",
        "eps_cons_low", "eps_cons_high", "eps_cons_n", "cons_source", "fy_label", "fy_eps_cons",
        "fy_eps_low", "fy_eps_high", "fy_eps_n", "fy_rev_cons_bn", "fy_rev_low_bn", "fy_rev_high_bn",
        "fy_rev_n", "fy_source", "updated"]


def real_date(v):
    """Wind returns 1899-12-30 for 'no date'."""
    return v.date() if isinstance(v, datetime) and v.year > 1900 else None


def num(v):
    return None if v is None or v != v else float(v)  # v != v catches NaN


def period_end(fy, q, fy_end_month):
    """Year-month a fiscal quarter ends in, e.g. NVDA (FY ends Jan) FY2027 Q3 -> 2026-10."""
    m = fy_end_month - 3 * (4 - q)
    y = fy
    while m <= 0:
        m += 12
        y -= 1
    return f"{y}-{m:02d}"


def wind_quarters(w, codes, last_fy):
    """{code: {(fy, q): {...}}} from Wind. Wind labels a fiscal quarter by rptDate: FY2027 Q3 is 20270930."""
    out = {c: {} for c in codes}
    fields = "stm_predict_issuingdate,stm_issuingdate,oper_rev,eps_diluted"
    for fy in range(FIRST_FY, last_fy + 1):
        for q, mmdd in enumerate(["0331", "0630", "0930", "1231"], start=1):
            d = w.wss(",".join(codes), fields, f"rptDate={fy}{mmdd};rptType=1;unit=1;currencyType=")
            if d.ErrorCode != 0:
                raise RuntimeError(f"Wind error {d.ErrorCode}: {d.Data}")
            for i, c in enumerate(codes):
                out[c][(fy, q)] = {
                    "expected": real_date(d.Data[0][i]), "reported": real_date(d.Data[1][i]),
                    "rev_ytd": num(d.Data[2][i]), "eps_ytd": num(d.Data[3][i]),
                }
    # Year-to-date to quarterly
    for qs in out.values():
        for (fy, q), r in qs.items():
            prev = qs.get((fy, q - 1)) if q > 1 else {"rev_ytd": 0.0, "eps_ytd": 0.0}
            r["rev"] = r["rev_ytd"] - prev["rev_ytd"] if prev and None not in (r["rev_ytd"], prev["rev_ytd"]) else None
            r["eps"] = r["eps_ytd"] - prev["eps_ytd"] if prev and None not in (r["eps_ytd"], prev["eps_ytd"]) else None
    return out


class Yahoo:
    def __init__(self):
        self.op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
        self.op.addheaders = [("User-Agent", UA)]
        try:
            self.op.open("https://fc.yahoo.com", timeout=15)  # sets the session cookie; answers 404
        except Exception:
            pass
        self.crumb = self.op.open("https://query1.finance.yahoo.com/v1/test/getcrumb", timeout=15).read().decode()

    def summary(self, ticker):
        url = (f"https://query2.finance.yahoo.com/v10/finance/quoteSummary/{urllib.parse.quote(ticker)}?"
               + urllib.parse.urlencode({"modules": "earningsTrend,earningsHistory,calendarEvents,financialData", "crumb": self.crumb}))
        return json.load(self.op.open(url, timeout=20))["quoteSummary"]["result"][0]


def nasdaq_next(ticker):
    """Next results date from Nasdaq.com (Zacks data): {date, confirmed, session, period_end}, or None.
    Zacks writes 'is expected* to report ... after market close' once the company has announced the date,
    and 'is estimated to report ...' while the date is its own estimate from past reporting dates."""
    req = urllib.request.Request(f"https://api.nasdaq.com/api/analyst/{urllib.parse.quote(ticker)}/earnings-date",
                                 headers={"User-Agent": UA, "Accept": "application/json"})
    text = (json.load(urllib.request.urlopen(req, timeout=20)).get("data") or {}).get("reportText") or ""
    m = re.search(r"is (expected\*|estimated) to report earnings on\s+(\d{2})/(\d{2})/(\d{4})"
                  r"(?:\s+(after market close|before market open))?", text)
    if not m:
        return None
    pe = re.search(r"fiscal Quarter ending (\w{3}) (\d{4})", text)
    return {"date": date(int(m[4]), int(m[2]), int(m[3])), "confirmed": m[1] == "expected*",
            "session": {"after market close": "AMC", "before market open": "BMO"}.get(m[5] or "", ""),
            "period_end": f"{pe[2]}-{MONTHS[pe[1]]:02d}" if pe and pe[1] in MONTHS else None}


def near(by_month, ym):
    """by_month[ym], else the month either side: Yahoo dates a 52/53-week quarter by month end, so Broadcom's
    quarter ending 2 Nov 2026 is 2026-10 there."""
    y, m = map(int, ym.split("-"))
    for k in (0, -1, 1):
        v = by_month.get(f"{y + (m + k - 1) // 12}-{(m + k - 1) % 12 + 1:02d}")
        if v:
            return v
    return None


def raw(d, k):
    v = (d or {}).get(k)
    return v.get("raw") if isinstance(v, dict) else None


def yahoo_ticker(code):
    return code.split(".")[0]


def r2(v, n=2):
    return "" if v is None else round(v, n)


def bn(v):
    return "" if not v else round(v / 1e9, 3)


def main():
    universe = [u for u in csv.DictReader(open(ROOT / "universe.csv", newline="", encoding="utf-8"))
                if u["type"] == "Equity Stock"]
    if not universe:
        print("No Equity Stock rows in universe.csv; calendar not built.")
        return
    from build_data import wind
    w = wind()
    codes = [u["code"] for u in universe]
    today = datetime.now(SGT).date()
    fy_end = dict(zip(codes, wind_ok(w.wss(",".join(codes), "fiscaldate", "")).Data[0]))  # e.g. '01-25'
    quarters = wind_quarters(w, codes, today.year + 1)

    try:
        yahoo = Yahoo()
    except Exception as e:
        yahoo = None
        print(f"Yahoo unavailable ({e}); consensus left blank")

    stamp = datetime.now(SGT).strftime("%Y-%m-%d %H:%M SGT")
    rows = []
    for u in universe:
        c = u["code"]
        fem = int(str(fy_end[c]).split("-")[0])
        ys = {}
        if yahoo:
            try:
                ys = yahoo.summary(yahoo_ticker(c))
            except Exception as e:
                print(f"Yahoo {c}: {e}")
        ccy = (ys.get("financialData") or {}).get("financialCurrency") or u["currency"]
        trends = [t for t in (ys.get("earningsTrend") or {}).get("trend", []) if t.get("endDate")]
        trend = {t["endDate"][:7]: t for t in trends if t.get("period") in ("0q", "+1q")}
        fy_trend = {t["endDate"][:7]: t for t in trends if t.get("period") in ("0y", "+1y")}
        hist = {datetime.fromtimestamp(h["quarter"]["raw"], timezone.utc).strftime("%Y-%m"): h
                for h in (ys.get("earningsHistory") or {}).get("history", []) if raw(h, "quarter")}
        ydates = [datetime.fromtimestamp(d["raw"], timezone.utc).date() for d in ((ys.get("calendarEvents") or {}).get("earnings") or {}).get("earningsDate", [])]
        try:
            nd = nasdaq_next(yahoo_ticker(c))
        except Exception as e:
            nd = None
            print(f"Nasdaq {c}: {e}")
        # The next report is the quarter after the last one Wind has a results date for
        done = [k for k, r in quarters[c].items() if r["reported"]]
        if done:
            fy, q = max(done)
            next_key = (fy + 1, 1) if q == 4 else (fy, q + 1)
        else:
            next_key = min((k for k, r in quarters[c].items() if r["expected"] and r["expected"] >= today), default=None)
        next_done = False
        for (fy, q), r in sorted(quarters[c].items()):
            pe = period_end(fy, q, fem)
            when = r["reported"] or r["expected"]
            date_source = "Wind" if r["reported"] else "Wind estimate"
            confirmed = r["reported"] is not None
            status_source = "Wind (reported)" if confirmed else ""
            session = u.get("earnings_session", "")
            if (fy, q) == next_key:
                # Wind first; a date the company has announced replaces Wind's estimate; Nasdaq's or Yahoo's
                # estimate fills in only when Wind has no date
                if nd and nd["confirmed"] and nd["period_end"] in (None, pe):
                    when, date_source = nd["date"], "Wind" if when == nd["date"] else "Nasdaq (Zacks)"
                    confirmed = True
                    session = session or nd["session"]
                elif not when and nd:
                    when, date_source = nd["date"], "Nasdaq (Zacks) estimate"
                elif not when and ydates and ydates[0] >= today:
                    when, date_source = ydates[0], "Yahoo estimate"
                status_source = "Nasdaq (Zacks)"
            if not when:
                continue
            reported = r["reported"] is not None and r["rev"] is not None
            row = {k: "" for k in COLS}
            row.update({
                "event_id": f"{yahoo_ticker(c)}-FY{fy % 100:02d}Q{q}", "code": c, "ticker": yahoo_ticker(c),
                "name": u["name"], "category": "Equity", "kind": "Earnings", "event_date": when.isoformat(),
                "session": session, "status": "Reported" if reported else "Expected",
                "date_status": "Confirmed" if confirmed else "Estimate", "date_status_source": status_source,
                "date_source": date_source,
                "fiscal_period": f"FY{fy % 100:02d} Q{q}", "period_end": pe, "currency": ccy,
                "rev_actual_bn": bn(r["rev"]) if reported else "",
                "eps_gaap": r2(r["eps"]) if reported and ccy == u["currency"] else "", "updated": stamp,
            })
            h = near(hist, pe)
            if h:
                act, est, surprise = raw(h, "epsActual"), raw(h, "epsEstimate"), raw(h, "surprisePercent")
                if act is not None and est and abs(est) >= 0.5 and abs(act) > 10 * abs(est):
                    # Not on one share basis, e.g. SK hynix's actual per ordinary share against an estimate per ADR
                    print(f"Yahoo {c} {pe}: actual EPS {act} vs estimate {est} not comparable; actual left out")
                    act = surprise = None
                row.update({"eps_cons": r2(est, 3), "eps_adj_actual": r2(act),
                            "eps_surprise_pct": r2(surprise * 100, 1) if surprise is not None else "",
                            "cons_source": "Yahoo"})
            t = near(trend, pe)
            if t and not reported:
                ee, re_ = t.get("earningsEstimate", {}), t.get("revenueEstimate", {})
                row.update({
                    "eps_cons": r2(raw(ee, "avg"), 3), "eps_cons_low": r2(raw(ee, "low")), "eps_cons_high": r2(raw(ee, "high")),
                    "eps_cons_n": raw(ee, "numberOfAnalysts") or "",
                    "rev_cons_bn": bn(raw(re_, "avg")), "rev_cons_low_bn": bn(raw(re_, "low")),
                    "rev_cons_high_bn": bn(raw(re_, "high")),
                    "rev_cons_n": raw(re_, "numberOfAnalysts") or "", "cons_source": "Yahoo",
                })
            if not reported:
                if not next_done and ydates and ydates[0] >= today and ydates[0] != when:
                    row["date_yahoo"] = ydates[0].isoformat()
                next_done = True
                ft = near(fy_trend, period_end(fy, 4, fem))
                if ft:
                    ee, re_ = ft.get("earningsEstimate", {}), ft.get("revenueEstimate", {})
                    row.update({
                        "fy_label": f"FY{fy % 100:02d}", "fy_eps_cons": r2(raw(ee, "avg"), 3),
                        "fy_eps_low": r2(raw(ee, "low")), "fy_eps_high": r2(raw(ee, "high")),
                        "fy_eps_n": raw(ee, "numberOfAnalysts") or "", "fy_rev_cons_bn": bn(raw(re_, "avg")),
                        "fy_rev_low_bn": bn(raw(re_, "low")), "fy_rev_high_bn": bn(raw(re_, "high")),
                        "fy_rev_n": raw(re_, "numberOfAnalysts") or "", "fy_source": "Yahoo",
                    })
            rows.append(row)

    rows.sort(key=lambda r: (r["event_date"], r["ticker"]))
    with open(OUT, "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, fieldnames=COLS)
        wr.writeheader()
        wr.writerows(rows)
    print(f"{stamp}: {len(rows)} earnings events -> {OUT.name}")


def wind_ok(d):
    if d.ErrorCode != 0:
        raise RuntimeError(f"Wind error {d.ErrorCode}: {d.Data}")
    return d


if __name__ == "__main__":
    main()
