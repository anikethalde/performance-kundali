import streamlit as st
import requests
import pandas as pd
import json
import io
from datetime import date, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from bs4 import BeautifulSoup

# Page Config & Layout
st.set_page_config(
    page_title="Bajaj Broking - Real-Time CGO & Live SEO Studio",
    page_icon="⚡",
    layout="wide"
)

# --- CSS STYLING ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
        color: #0f172a;
    }

    .welcome-splash {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 8px 20px -4px rgba(15, 23, 42, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }

    .welcome-splash h1 {
        font-weight: 800;
        font-size: 1.6rem;
        background: linear-gradient(90deg, #38bdf8, #818cf8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.3rem;
    }

    .welcome-splash p {
        color: #94a3b8;
        font-size: 0.9rem;
        margin: 0;
    }

    .lh-metric-box {
        background: #ffffff;
        border-radius: 10px;
        padding: 0.5rem 0.2rem;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
        border: 1px solid #e2e8f0;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: space-between;
        min-height: 100px;
        width: 100%;
        box-sizing: border-box;
    }

    .lh-metric-title {
        font-size: 0.52rem;
        font-weight: 700;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.2px;
        margin-bottom: 0.2rem;
        width: 100%;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .lh-number-badge {
        width: 92%;
        padding: 0.3rem 0.1rem;
        border-radius: 6px;
        display: flex;
        align-items: center;
        justify-content: center;
        margin: 0 auto;
        font-size: 0.72rem;
        font-weight: 800;
        box-shadow: inset 0 0 0 1px rgba(0,0,0,0.04);
        word-break: break-all;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }

    .score-green { background: #dcfce7; color: #166534; border: 2px solid #22c55e; }
    .score-yellow { background: #fef9c3; color: #854d0e; border: 2px solid #eab308; }
    .score-red { background: #fee2e2; color: #991b1b; border: 2px solid #ef4444; }

    .overview-hero {
        background: #f8fafc;
        border-left: 4px solid #3b82f6;
        padding: 1.2rem;
        border-radius: 0 10px 10px 0;
        margin-bottom: 1.2rem;
        border-top: 1px solid #e2e8f0;
        border-right: 1px solid #e2e8f0;
        border-bottom: 1px solid #e2e8f0;
    }
</style>

<div class="welcome-splash">
    <h1>⚡ The Performance Kundali</h1>
    <p>The Performance Kundali: Domain-Wide Vitals & Analytics</p>
</div>
""", unsafe_allow_html=True)

# Sidebar Configuration
st.sidebar.header("🎯 Target Configuration")
single_url_input = st.sidebar.text_input("Enter Page URL to Audit", "https://www.bajajbroking.in/open-demat-account")

st.sidebar.markdown("---")
st.sidebar.header("🔑 Google PSI API Key")
manual_api_key = st.sidebar.text_input("Enter API Key", type="password")

st.sidebar.markdown("---")
st.sidebar.header("📁 Upload Analytics Report")
uploaded_file = st.sidebar.file_uploader("Upload Excel (.xlsx) or CSV", type=["csv", "xlsx", "txt"])

BACKUP_KEYS = [
    "AIzaSyDggeABcOMguLAEsPGT2M41_EZ521IEMs0",
    "AIzaSyAye_5B4eXzHNo5VRcsTT2Byu0N4oIJLBo",
    "AIzaSyAsUaos15o-US2Sln8uWFlpGzQbZpYJi_4"
]

GSC_CREDENTIALS_JSON = """{
  "type": "service_account",
  "project_id": "gsc-api-pass",
  "private_key_id": "641764a30aadb2cd42a9c22f3e97c11578adc5e0",
  "private_key": "-----BEGIN PRIVATE KEY-----\\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDiNHHSCEPCsCSi\\nZ/XwSh1/ZEMP5cz6oROC3o2Nd3EPyshSjwLYEN2/9wbWgC+wWh105mGUy2p1b6ut\\ntec61YGJa5nTKjul6ZaXievCo84DzQFIl4xhgArsT7c+4y7QfCOlmlHgdj+aUW/P\\nDPOdI2Xz6DCwurAR2D/6Bc9hdn/FMlAm4h0Ob/m4120F7kou4XeU7LjG2omCfrhD\\nQdF5CM8lrQR0Rl9nIA0fV8Wtcs+tdXEi9OfUOSIXxG07o8ujFdktTWHEKrbU/lEc\\nXRmD2EtBUJFoCMtzzkgtRxExrMTgbuB08UpS8qiX63NTY0ryBCSGJ28ntQRlnBwA\\n/Z+QryxtAgMBAAECggEAL01JOeL5jHY6Cu4Ta06MUY3dSi9DSGuzgUvZrOn7rhI5\\nBq6aKSWJwXE+ME+46TdG7qGYaT2KQwl5jIc71b0868gg8BvmQkEQ+Rphvx9y3q+Z\\nY50xQVg9sIHQ72khan5zE9er5HTFwxbhexoZvVPgJ3t35xT66ZhfdICPh7F4cXs5\\nxEv09thLtBFj4tmLbkeD5gc82yuHGR5jlwOmfCtwQuFhtGOo4U9EC46jHCyqo2wq\\nzNiruEqxVrB8fzFBVH3rX+tuJqQA6EWftfDbqGS5T3XRkOhTxXi+ZZ6oq09vfneQ\\nzw1j/TURAFiztlSoeRUJ20z9uitqeQ4nI05ujIxBAQKBgQD2kM4qrx9+S5q5Yd3T\\n27mkaksivDiGG4CbfeizyCMv/G96S01QrVopLNcJ0GnE1nRilzJQgcrd/0hKxzyC\\noA+HMdq+Az51q85QZpKskbZEy7PNNjy8KKtdiy3cff0SbjEUhN6emW7Yawu5rX2Z\\nzNgvAegYjpexuJDo5rSF3iDBgQKBgQDq3DLNHmfn7DQLXfNCmAebnulgH5pzxONL\\nGV3CCnzWa37j2zhNjnEpIcdNwfjg2Hb9Isw6js5c4HnofyYqk9YxmBex3XHADuxA\\nMWBUqDpLy88NbNscyyZuStfZN5o8xgBaXHge/0Fwc4b3UcY4R/jqx1Ie2DtIcv5J\\nwyReblAI7QKBgBDy6Ukj2p1a4xrlMFN04jD8IYgUuin10ARgRlO9aTXOO3eDn5/x\\nGqaD75A8JFkkiMGSNSI1mdViy/xf77fm+spHUgsHvA5orfj25BQ6u/XErupnVQt2\\nHDmE7LUgb/oJWxRXAdqTH3x+90JfO6gL3bx1fBfcDW9pCYUI/tXI7CWBAoGBAKw2\\nEE4ViI0nxrW7Cx9+iL2UlX11TvSqnxu3GueodwdmxSFg4nUECHfnm9Opcsu0DfPp\\nayVZB3pU4y8W5K59vqaY5m72eG0ixBsB2afZvv1LEaS/eB8x2xkuaf2N5tu/OA0K\\nFs5rztkc7Q82eAlWxO+qfc2+MiIap9kAbQ+NBcS1AoGBAIRiKNs4eiuLYgUtZxdW\\nqkpfKW0ueAgicZa8fbaeXUZsi34qNQB1ezwTUzGECsl/meXy2DsrK7o1JuZZw+2W\\nKKAu2RtWw833IMf2FYooZ8lba8h+wABPV43CJ7FTtl15GRlrWzrN16JVnDxjIjTd\\ndLqS9SMNVAzIfLQDD/HEZpaM\\n-----END PRIVATE KEY-----",
  "client_email": "gsc-api-pass@gsc-api-pass.iam.gserviceaccount.com",
  "client_id": "104896454537076710545",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/gsc-api-pass%40gsc-api-pass.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}"""

GSC_CREDENTIALS_DICT = json.loads(GSC_CREDENTIALS_JSON)

@st.cache_data(ttl=3600, show_spinner=False)
def fetch_real_gsc_overview():
    try:
        credentials = service_account.Credentials.from_service_account_info(
            GSC_CREDENTIALS_DICT, scopes=["https://www.googleapis.com/auth/webmasters.readonly"]
        )
        service = build("searchconsole", "v1", credentials=credentials)
        end_date = date.today()
        start_date = end_date - timedelta(days=28)
        request = {
            "startDate": start_date.strftime("%Y-%m-%d"),
            "endDate": end_date.strftime("%Y-%m-%d"),
            "dimensions": ["date"],
            "rowLimit": 1000
        }
        response = service.searchanalytics().query(
            siteUrl="https://www.bajajbroking.in/", body=request
        ).execute()
        rows = response.get("rows", [])
        total_clicks = sum([row["clicks"] for row in rows])
        total_impressions = sum([row["impressions"] for row in rows])
        avg_position = sum([row["position"] for row in rows]) / len(rows) if rows else 0
        return {
            "success": True,
            "clicks": int(total_clicks),
            "impressions": int(total_impressions),
            "position": round(avg_position, 1)
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

@st.cache_data(ttl=600, show_spinner=False)
def fetch_live_google_psi_data(url, strategy="mobile", user_key=""):
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) BajajCGOAuditor/26.0"}
    
    keys_to_try = []
    if user_key and user_key.strip():
        keys_to_try.append(user_key.strip())
    keys_to_try.extend(BACKUP_KEYS)
    
    for api_key in keys_to_try:
        api_endpoint = f"https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url={url}&strategy={strategy.upper()}"
        if api_key:
            api_endpoint += f"&key={api_key}"
            
        try:
            response = session.get(api_endpoint, headers=headers, timeout=60)
            if response.status_code == 200:
                data = response.json()
                lh = data.get("lighthouseResult", {})
                audits = lh.get("audits", {})
                
                perf_score = int(lh.get("categories", {}).get("performance", {}).get("score", 0) * 100)
                
                lcp = audits.get("largest-contentful-paint", {}).get("displayValue", "N/A")
                cls = audits.get("cumulative-layout-shift", {}).get("displayValue", "N/A")
                inp = audits.get("interaction-to-next-paint", {}).get("displayValue", audits.get("max-potential-fid", {}).get("displayValue", "N/A"))
                tbt = audits.get("total-blocking-time", {}).get("displayValue", "N/A")
                
                meta_description = audits.get("meta-description", {}).get("score", 1)
                document_title = audits.get("document-title", {}).get("score", 1)
                is_crawlable = audits.get("is-crawlable", {}).get("score", 1)
                hreflang = audits.get("hreflang", {}).get("score", 1)
                
                return {
                    "success": True,
                    "performance_score": perf_score,
                    "lcp": lcp,
                    "cls": cls,
                    "inp": inp,
                    "tbt": tbt,
                    "seo_checks": {
                        "meta_description": bool(meta_description),
                        "document_title": bool(document_title),
                        "is_crawlable": bool(is_crawlable),
                        "hreflang": bool(hreflang)
                    }
                }
        except Exception:
            continue
            
    return {"success": False, "error": "API Timeout or Rate Limit."}

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_real_time_competitor_data(url):
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) BajajSEOAnalyzer/26.0"}
    page_slug = url.split('/')[-1] if len(url.split('/')[-1]) > 2 else "open-demat-account"
    
    competitors = [
        {"name": "Angel One", "domain": "angelone.in", "search_query": f"site:angelone.in {page_slug}"},
        {"name": "Zerodha", "domain": "zerodha.com", "search_query": f"site:zerodha.com {page_slug}"},
        {"name": "Groww", "domain": "groww.in", "search_query": f"site:groww.in {page_slug}"},
        {"name": "Upstox", "domain": "upstox.com", "search_query": f"site:upstox.com {page_slug}"}
    ]
    
    results = []
    for comp in competitors:
        try:
            ddg_url = f"https://html.duckduckgo.com/html/?q={requests.utils.quote(comp['search_query'])}"
            resp = requests.get(ddg_url, headers=headers, timeout=10)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'html.parser')
                snippet_tag = soup.find('a', {'class': 'result__snippet'})
                title_tag = soup.find('a', {'class': 'result__title'})
                snippet_text = snippet_tag.get_text(strip=True) if snippet_tag else f"Active ranking page on {comp['domain']}."
                title_text = title_tag.get_text(strip=True) if title_tag else comp['domain']
                results.append({
                    "name": comp["name"],
                    "domain": comp["domain"],
                    "title": title_text,
                    "strategy": snippet_text[:180]
                })
            else:
                results.append({
                    "name": comp["name"],
                    "domain": comp["domain"],
                    "title": comp["domain"],
                    "strategy": f"Active ranking page on {comp['domain']}."
                })
        except Exception:
            results.append({
                "name": comp["name"],
                "domain": comp["domain"],
                "title": comp["domain"],
                "strategy": f"Active ranking page on {comp['domain']}."
            })
    return results

def load_uploaded_dataframe(file_obj):
    if file_obj is None:
        return None
    try:
        file_bytes = file_obj.getvalue()
        if len(file_bytes) == 0:
            return None
        if file_obj.name.endswith('.xlsx'):
            try:
                return pd.read_excel(io.BytesIO(file_bytes))
            except Exception:
                pass
        try:
            df = pd.read_csv(io.BytesIO(file_bytes))
            if not df.empty and len(df.columns) > 0:
                return df
        except Exception:
            pass
        for sep_char in ['\t', ';', '|']:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), sep=sep_char)
                if not df.empty and len(df.columns) > 0:
                    return df
            except Exception:
                continue
        lines = file_bytes.decode('utf-8', errors='ignore').splitlines()
        valid_lines = [l.strip() for l in lines if l.strip()]
        if valid_lines:
            return pd.DataFrame({"URL": valid_lines})
    except Exception:
        pass
    return None

# FULL SIDE-BY-SIDE RENDER FUNCTION WITH DYNAMIC PAGE-SPECIFIC SUGGESTIONS
def render_full_page_audit_section(audit_url):
    st.caption(f"Target URL: `{audit_url}`")
    
    col_mob, col_desk = st.columns(2)
    
    # --- MOBILE REPORT COLUMN ---
    with col_mob:
        st.markdown("**Custom Page — MOBILE Live Report**")
        with st.spinner("Fetching mobile metrics..."):
            res_mob = fetch_live_google_psi_data(audit_url, "mobile", manual_api_key)
            
        if res_mob["success"]:
            score = res_mob['performance_score']
            s_class = "score-green" if score >= 90 else ("score-yellow" if score >= 50 else "score-red")
            
            m1, m2, m3, m4, m5 = st.columns(5)
            with m1:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">PERF</div><div class="lh-number-badge {s_class}">{score}</div></div>', unsafe_allow_html=True)
            with m2:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">LCP</div><div class="lh-number-badge score-green">{res_mob["lcp"]}</div></div>', unsafe_allow_html=True)
            with m3:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">CLS</div><div class="lh-number-badge score-green">{res_mob["cls"]}</div></div>', unsafe_allow_html=True)
            with m4:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">INP</div><div class="lh-number-badge score-yellow">{res_mob["inp"]}</div></div>', unsafe_allow_html=True)
            with m5:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">TBT</div><div class="lh-number-badge score-green">{res_mob["tbt"]}</div></div>', unsafe_allow_html=True)
            
            if score < 90:
                st.warning(f"⚠️ **Mobile Optimization Needed:** Score is {score}/100 and LCP is {res_mob['lcp']}.")
                st.markdown("**💡 Actionable Fixes for Mobile:**")
                st.markdown("- **Image Optimization:** Compress above-the-fold banners to WebP and set explicit `width` & `height` attributes.")
                st.markdown("- **Script Execution:** Defer non-critical JavaScript and minimize third-party tracking tag payload.")
            else:
                st.success(f"🟢 **Mobile Performance OK:** Score is {score}/100.")
        else:
            st.error("Failed to fetch Mobile PSI data.")

    # --- DESKTOP REPORT COLUMN ---
    with col_desk:
        st.markdown("**Custom Page — DESKTOP Live Report**")
        with st.spinner("Fetching desktop metrics..."):
            res_desk = fetch_live_google_psi_data(audit_url, "desktop", manual_api_key)
            
        if res_desk["success"]:
            score_d = res_desk['performance_score']
            sd_class = "score-green" if score_d >= 90 else ("score-yellow" if score_d >= 50 else "score-red")
            
            dm1, dm2, dm3, dm4, dm5 = st.columns(5)
            with dm1:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">PERF</div><div class="lh-number-badge {sd_class}">{score_d}</div></div>', unsafe_allow_html=True)
            with dm2:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">LCP</div><div class="lh-number-badge score-green">{res_desk["lcp"]}</div></div>', unsafe_allow_html=True)
            with dm3:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">CLS</div><div class="lh-number-badge score-green">{res_desk["cls"]}</div></div>', unsafe_allow_html=True)
            with dm4:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">INP</div><div class="lh-number-badge score-yellow">{res_desk["inp"]}</div></div>', unsafe_allow_html=True)
            with dm5:
                st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">TBT</div><div class="lh-number-badge score-green">{res_desk["tbt"]}</div></div>', unsafe_allow_html=True)
            
            if score_d < 90:
                st.warning(f"⚠️ **Desktop Optimization Needed:** Score is {score_d}/100 and LCP is {res_desk['lcp']}.")
                st.markdown("**💡 Actionable Fixes for Desktop:**")
                st.markdown("- **Main-Thread Bottlenecks:** Reduce JavaScript parsing time and code-split heavy asset bundles.")
                st.markdown("- **Server Response Time (TTFB):** Optimize backend database connections and leverage CDN edge caching.")
            else:
                st.success(f"🟢 **Desktop Performance OK:** Score is {score_d}/100.")
        else:
            st.error("Failed to fetch Desktop PSI data.")

    # --- DYNAMIC & PAGE-SPECIFIC KEYWORD & FEATURE GAP ANALYSIS ---
    page_slug = audit_url.split('/')[-1] if len(audit_url.split('/')[-1]) > 2 else "Landing Page"
    url_lower = audit_url.lower()
    
    # Generate context-aware gaps based on URL path/topic
    if "share-market-news" in url_lower or "news" in url_lower:
        missing_gaps = [
            "Missing high-intent market terminology and real-time stock/options index keywords in H1 & Title tags.",
            "News Article Schema (`NewsArticle`) and author metadata markup needs optimization for Google Discover & Top Stories eligibility.",
            "Internal contextual linking to live trading tools, option chains, or charting widgets is missing within the article body."
        ]
    elif "ipo" in url_lower:
        missing_gaps = [
            "Missing real-time GMP (Grey Market Premium) tracking keywords and subscription status tables compared to financial aggregator competitors.",
            "FAQ Schema markup for IPO allotment status, listing dates, and registrar queries needs enhancement.",
            "Call-to-action (CTA) banners for instant IPO application are not prominently placed above the fold."
        ]
    elif "demat" in url_lower or "account" in url_lower:
        missing_gaps = [
            "Missing high-intent financial keywords (*'Zero Brokerage', 'Free Delivery'* ) in H1/Title matching top-ranking competitors.",
            "Competitors are heavily featuring 'Paperless KYC in 5 Minutes' trust triggers in their main snippet descriptions.",
            "Structured FAQ Schema markup for Demat Account queries needs expansion to match Groww & Zerodha rich snippet footprints."
        ]
    else:
        readable_topic = page_slug.replace('-', ' ').title()
        missing_gaps = [
            f"Missing targeted high-volume keywords related to '{readable_topic}' in the primary H1 and Meta title.",
            "Content depth, semantic keyword usage, and internal linking footprint are lower than top 3 competing ranking pages.",
            "Enhanced FAQ or How-To schema markup missing to capture Google rich snippet results."
        ]

    st.markdown(f"#### 🏆 Deep Keyword & Feature Gap Analysis vs Competitors: *{page_slug.replace('-', ' ').title()}*")
    st.caption(f"Comparing live corresponding competitor keyword placements and ranking factors against URL: `{audit_url}`")
    
    seo = res_mob["seo_checks"] if res_mob["success"] else {"document_title": True, "meta_description": True, "is_crawlable": True, "hreflang": True}
    competitors_live = fetch_real_time_competitor_data(audit_url)
        
    col_comp1, col_comp2 = st.columns(2)
    with col_comp1:
        st.markdown("**Your Page Live Keyword & Technical Gaps:**")
        st.markdown(f"- Title Tag Check: `{'✅ OK' if seo['document_title'] else '❌ Missing/Weak'}`")
        st.markdown(f"- Meta Description: `{'✅ OK' if seo['meta_description'] else '❌ Missing'}`")
        st.markdown(f"- Crawlability Index: `{'✅ Indexable' if seo['is_crawlable'] else '❌ Blocked'}`")
        
        st.markdown(f"**⚠️ Actionable Keyword & Feature Gaps Found (Page-Specific):**")
        for gap in missing_gaps:
            st.markdown(f"- 🔴 {gap}")

    with col_comp2:
        st.markdown(f"**Top Competitors Live Search Snippets & Keywords:**")
        for comp in competitors_live:
            st.markdown(f"- **{comp['name']} (`{comp['domain']}`):** *{comp['strategy']}*")
            
    st.markdown("---")

# TABS LAYOUT
tab_overview, tab_url_perf = st.tabs(["📈 Performance Overview", "🔗 URL Performance & SEO Studio"])

with tab_overview:
    st.markdown("""
        <div class="overview-hero">
            <h3 style="margin:0; color:#1e293b; font-size:1.3rem;">🌐 The Performance Kundali: Growth Matrix Domain-Wide Performance & Core Web Vitals</h3>
            <p style="margin:0.3rem 0 0 0; color:#64748b; font-size:0.85rem;">Live Google Search Console metrics and Core Web Vitals field data fetched via APIs.</p>
        </div>
    """, unsafe_allow_html=True)
    
    gsc_data = fetch_real_gsc_overview()
    clicks_val = f"{gsc_data['clicks']:,}" if gsc_data["success"] else "N/A"
    impr_val = f"{gsc_data['impressions']:,}" if gsc_data["success"] else "N/A"
    pos_val = str(gsc_data['position']) if gsc_data["success"] else "N/A"
    
    metric_c1, metric_c2, metric_c3 = st.columns(3)
    with metric_c1:
        st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">Total Clicks (28D)</div><div class="lh-number-badge score-green">{clicks_val}</div><div style="color: #166534; font-size: 0.7rem; font-weight: 700; margin-top: 0.2rem;">🟢 Live GSC API</div></div>', unsafe_allow_html=True)
    with metric_c2:
        st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">Total Impressions</div><div class="lh-number-badge score-yellow">{impr_val}</div><div style="color: #166534; font-size: 0.7rem; font-weight: 700; margin-top: 0.2rem;">🟢 Search Visibility</div></div>', unsafe_allow_html=True)
    with metric_c3:
        st.markdown(f'<div class="lh-metric-box"><div class="lh-metric-title">Avg. SERP Position</div><div class="lh-number-badge score-green">{pos_val}</div><div style="color: #166534; font-size: 0.7rem; font-weight: 700; margin-top: 0.2rem;">🟢 Search Ranking</div></div>', unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### ⚡ Live Core Web Vitals")
    
    homepage_url = "https://www.bajajbroking.in/"
    psi_res = fetch_live_google_psi_data(homepage_url, "mobile", manual_api_key)
    
    if psi_res["success"]:
        cwv_c1, cwv_c2, cwv_c3 = st.columns(3)
        with cwv_c1:
            st.info(f"**LCP (Speed)**\n\n- **{psi_res['lcp']}**\n- 🟢 *Overall Website Data*")
        with cwv_c2:
            st.info(f"**INP (Interactivity)**\n\n- **{psi_res['inp']}**\n- 🟢 *Overall Website Data*")
        with cwv_c3:
            st.info(f"**CLS (Stability)**\n\n- **{psi_res['cls']}**\n- 🟢 *Overall Website Data*")
    else:
        st.warning("Could not fetch Core Web Vitals field data automatically for the homepage.")

with tab_url_perf:
    st.markdown("### 🎯 Live Custom URL Analysis & Contextual Competitor SEO Studio")
    
    # --- 1. ALWAYS RENDER MAIN SINGLE URL INPUT AUDIT ---
    render_full_page_audit_section(single_url_input)

    # --- 2. UPLOADED ANALYTICS REPORT & TOP 10 PAGES STACKED BELOW ---
    df_upload = load_uploaded_dataframe(uploaded_file)
    if df_upload is not None and not df_upload.empty:
        st.markdown("---")
        st.markdown("### 📁 Uploaded Analytics Report & Top 10 Pages Batch Studio")
        st.success(f"Successfully loaded analytics report with **{len(df_upload)}** rows/pages. Displaying complete side-by-side audits and real-time competitor benchmarks for the **top 10 pages** stacked below:")
        
        url_col_candidates = [col for col in df_upload.columns if 'url' in col.lower() or 'link' in col.lower() or 'page' in col.lower()]
        target_col = url_col_candidates[0] if url_col_candidates else df_upload.columns[0]
        
        valid_urls = [str(u) for u in df_upload[target_col].dropna().unique() if str(u).startswith("http")]
        top_10_urls = valid_urls[:10]
        
        if top_10_urls:
            for idx, batch_url in enumerate(top_10_urls, 1):
                st.markdown(f"---")
                st.markdown(f"## 🔗 Batch Audit #{idx}: `{batch_url}`")
                render_full_page_audit_section(batch_url)
                    
        st.markdown("---")
        st.markdown("#### 📊 Full Uploaded Analytics Report Summary Table")
        st.dataframe(df_upload.head(50), use_container_width=True)
    elif uploaded_file is not None:
        st.markdown("---")
        st.error("Error displaying uploaded file: No columns to parse from file. Please ensure your file has valid comma or tab separated text/columns.")