"""Charmigos CRM app."""
import io, os, warnings, hashlib
from datetime import timedelta, datetime
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from scipy import stats
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import MinMaxScaler
warnings.filterwarnings("ignore")
try:
    from preprocessing import preprocess_shopee, preprocess_tiktok, merge_to_gabungan
    from storage import save_raw_file, save_gabungan, load_gabungan_bytes, save_strategy_doc, list_strategy_docs, sync_hasil_background, GABUNGAN_PATH, BASE_DIR, PROCESSED_DIR
    _MODULES_OK = True
except ImportError:
    _MODULES_OK = False
    BASE_DIR = Path(".")
    PROCESSED_DIR = BASE_DIR / "Data" / "Data hasil pengolahan"
    GABUNGAN_PATH = PROCESSED_DIR / "Transaksi_Gabungan.xlsx"
st.set_page_config(page_title="Charmigos CRM", page_icon="🛍️", layout="wide", initial_sidebar_state="expanded")
LOGO_PATH = Path(__file__).parent / "Gambar" / "logo charmigos.png"
USERS = {"manager": {"pwd": hashlib.sha256("charmigos123".encode()).hexdigest(), "role":"Manager","name":"Manager Charmigos"}, "staff": {"pwd": hashlib.sha256("staff123".encode()).hexdigest(), "role":"Staff","name":"Staff Operasional"}}
B2B_THRESHOLD = 300_000
KANAL_COLOR = {"Shopee":"#EE4D2D","TikTok":"#010101","WhatsApp":"#25D366"}
SEG_ORDER = ["Champions","Loyal Customers","Potential Loyalist","Recent Customers","Promising","Customers Needing Attention","About To Sleep","At Risk","Can't Lose Them","Hibernating","Lost"]
SEG_COLOR = {"Champions":"#16a34a","Loyal Customers":"#22c55e","Potential Loyalist":"#86efac","Recent Customers":"#3b82f6","Promising":"#93c5fd","Customers Needing Attention":"#f59e0b","About To Sleep":"#fcd34d","At Risk":"#ef4444","Can't Lose Them":"#b91c1c","Hibernating":"#94a3b8","Lost":"#6b7280"}
STRATEGIES = {"Champions":"Beri imbalan, jadikan pengadopsi awal & pelaku promosi.","Loyal Customers":"Retensi melalui retention voucher 10%, post-purchase feedback, co-creation engagement, restock and new product notification, serta personalized product recommendation.","Potential Loyalist":"Program membership, cross-sell, personalisasi rekomendasi.","Recent Customers":"Onboarding, welcome offer, bangun hubungan awal.","Promising":"Brand awareness, uji coba gratis, penawaran pertama.","Customers Needing Attention":"Penawaran waktu terbatas, reaktivasi berbasis histori.","About To Sleep":"Produk populer + diskon, rebuild relasi.","Can't Lose Them":"Hubungi langsung, update produk, jangan kehilangan ke kompetitor.","At Risk":"Reaktivasi melalui reactivation voucher 20%, penyampaian product update, dan personalized product recommendation.","Hibernating":"Diskon khusus, produk relevan, rebuild brand value.","Lost":"Reaktivasi dengan reward besar atau pertimbangkan melepas."}
PRIORITY = {"Champions":"high","Loyal Customers":"high","Potential Loyalist":"medium","Recent Customers":"medium","Promising":"medium","Customers Needing Attention":"medium","About To Sleep":"low","At Risk":"high","Can't Lose Them":"high","Hibernating":"low","Lost":"low"}
SEG_DETAIL = {"Champions":{"R":"4–5","FM":"4–5","desc":"Pelanggan terbaik: sering, baru, nilai tinggi.","strategi":"Reward, brand ambassador, early adopter."},"Loyal Customers":{"R":"2–5","FM":"3–5","desc":"Beli konsisten, mudah di-upsell.","strategi":"Retention voucher 10%, post-purchase feedback, co-creation engagement, restock and new product notification, serta personalized product recommendation."},"Potential Loyalist":{"R":"3–5","FM":"1–3","desc":"Baru, berpotensi loyal.","strategi":"Membership, cross-sell."},"Recent Customers":{"R":"4–5","FM":"0–1","desc":"Sangat baru, belum beli banyak.","strategi":"Onboarding, welcome offer."},"Promising":{"R":"3–4","FM":"0–1","desc":"Baru, engagement rendah.","strategi":"Brand awareness, trial gratis."},"Customers Needing Attention":{"R":"2–3","FM":"2–3","desc":"Aktif menurun.","strategi":"Penawaran terbatas, reaktivasi."},"About To Sleep":{"R":"2–3","FM":"0–2","desc":"Frekuensi & nilai rendah.","strategi":"Produk populer + diskon."},"At Risk":{"R":"0–2","FM":"2–5","desc":"Dulu aktif, sekarang hilang.","strategi":"Reactivation voucher 20%, product update, dan personalized product recommendation."},"Can't Lose Them":{"R":"0–1","FM":"4–5","desc":"Bernilai tinggi, hilang.","strategi":"Hubungi langsung, jangan kalah ke kompetitor."},"Hibernating":{"R":"1–2","FM":"1–2","desc":"Tidak aktif lama, nilai rendah.","strategi":"Diskon khusus, rebuild value."},"Lost":{"R":"0–1","FM":"0–1","desc":"Sangat lama tidak beli.","strategi":"Reaktivasi besar atau lepas."}}
STRATEGY_CHECKLISTS = {"At Risk":[("voucher_20","Reactivation Voucher 20%"),("product_update","Product Update"),("personalized_recommendation","Personalized Product Recommendation")],"Loyal Customers":[("voucher_10","Retention Voucher 10%"),("post_purchase_feedback","Post-purchase Feedback"),("co_creation","Co-creation Engagement"),("restock_notification","Restock & New Product Notification"),("personalized_recommendation","Personalized Product Recommendation")]}
REQUIRED_COLS = ["No. Pesanan","Waktu Pemesanan","Username (Pembeli)","User ID","Total Harga Produk","Kanal"]
RFM_FEATS = ["Recency","Frequency","Monetary"]
BOXCOX_FEATS = ["Recency","Monetary"]
RFM_SCORE_LABELS = [1,2,3,4,5]
def _check(u, p):
    usr = USERS.get(u.lower())
    return usr if usr and usr["pwd"] == hashlib.sha256(p.encode()).hexdigest() else None
def _seg(r, fm):
    r,fm = round(r),round(fm)
    if r>=4 and fm>=4: return "Champions"
    if r>=2 and fm>=3: return "Loyal Customers"
    if r>=3 and 1<=fm<=3: return "Potential Loyalist"
    if r>=4 and fm<=1: return "Recent Customers"
    if (r==3 or r==4) and fm<=1: return "Promising"
    if (r==2 or r==3) and (fm==2 or fm==3): return "Customers Needing Attention"
    if (r==2 or r==3) and fm<=2: return "About To Sleep"
    if r<=1 and fm>=4: return "Can't Lose Them"
    if r<=2 and fm>=2: return "At Risk"
    if (r==1 or r==2) and (fm==1 or fm==2): return "Hibernating"
    if r<=1 and fm<=1: return "Lost"
    return "Uncategorized"
def to_xl(df, sheet="Sheet1", text_cols=None):
    text_cols=[c for c in (text_cols or []) if c in df.columns]
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="xlsxwriter") as w:
        d=df.copy()
        for c in text_cols:
            d[c]=d[c].apply(lambda v:"" if pd.isna(v) else str(int(v)) if isinstance(v,(float,np.floating)) and float(v).is_integer() else str(v))
        d.to_excel(w,index=False,sheet_name=sheet)
        wb,ws=w.book,w.sheets[sheet]
        fmt=wb.add_format({"num_format":"@"})
        for i,c in enumerate(d.columns):
            if c in text_cols: ws.set_column(i,i,max(len(c)+2,18),fmt)
    return buf.getvalue()
def add_log(action,detail=""):
    if "log" not in st.session_state: st.session_state.log=[]
    st.session_state.log.append({"Waktu":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"Pengguna":st.session_state.get("uname","—"),"Role":st.session_state.get("urole","—"),"Aksi":action,"Detail":detail})
def maybe_sync_download(clicked, df, filename):
    if clicked and _MODULES_OK: sync_hasil_background(df, filename)
def get_quintile_ranges_ranked(series):
    ranked = series.rank(method="first"); quantiles = np.quantile(ranked.dropna(), [0,0.2,0.4,0.6,0.8,1.0]); ranges = []
    for i in range(5): ranges.append((series[ranked >= quantiles[i]].min(), series[ranked <= quantiles[i+1]].max()))
    return ranges
def relabel_clusters_by_recency(df, cluster_col="Cluster", recency_col="Recency"):
    order = df.groupby(cluster_col)[recency_col].mean().sort_values(ascending=False).index.tolist()
    return df[cluster_col].map({old:new for new,old in enumerate(order, start=1)}).astype(int)
def build_rfm(df, snap, extra_aggs=None):
    aggs = {"Username":("Username (Pembeli)","first"),"Recency":("Waktu Pemesanan",lambda x:(snap - x.max()).days),"Frequency":("No. Pesanan","nunique"),"Monetary":("Total Harga Produk","sum")}
    if extra_aggs: aggs.update(extra_aggs)
    return df.sort_values("Waktu Pemesanan").groupby("User ID").agg(**aggs).reset_index()
def _build_cluster_profile(rfm_b2c):
    cl=rfm_b2c.groupby("Cluster").agg(Recency=("Recency","mean"),Frequency=("Frequency","mean"),Monetary=("Monetary","mean"),R_norm=("R_norm","mean"),F_norm=("F_norm","mean"),M_norm=("M_norm","mean"),R_score=("R_score","mean"),F_score=("F_score","mean"),M_score=("M_score","mean"),FM_score=("FM_score","mean"),Count=("User ID","count")).reset_index().sort_values("Cluster").reset_index(drop=True)
    for c in ["R_score","F_score","M_score"]: cl[c]=np.floor(cl[c]).astype(int).clip(1,5)
    cl["FM_score"]=np.floor(cl["FM_score"]).astype(int).clip(1,5); cl["Segment"]=cl.apply(lambda x:_seg(x["R_score"],x["FM_score"]),axis=1); cl["Strategi"]=cl["Segment"].map(STRATEGIES)
    return cl
def prepare_b2c_clustering(rfm_b2c, k_range, safe_boxcox=False, include_ranges=False):
    d_raw = rfm_b2c[RFM_FEATS].copy(); bc = d_raw.copy(); lams = {}
    for col in BOXCOX_FEATS:
        try: bc[col], lams[col] = stats.boxcox(bc[col] + 1)
        except Exception:
            if not safe_boxcox: raise
            lams[col] = None
    d_norm = pd.DataFrame(MinMaxScaler().fit_transform(bc[RFM_FEATS]), columns=RFM_FEATS, index=rfm_b2c.index)
    rfm_b2c = rfm_b2c.copy(); rfm_b2c[["R_norm","F_norm","M_norm"]] = d_norm.values
    X, ine, sil = rfm_b2c[["R_norm","F_norm","M_norm"]].values, [], []
    for kk in k_range:
        km = KMeans(n_clusters=kk, init="k-means++", n_init=10, random_state=42)
        lb = km.fit_predict(X); ine.append(km.inertia_); sil.append(silhouette_score(X, lb, sample_size=min(5000, len(X)), random_state=42))
    bk = list(k_range)[sil.index(max(sil))]; raw_cluster = KMeans(n_clusters=bk, init="k-means++", n_init=20, random_state=42).fit_predict(X)
    rfm_b2c["Cluster"] = relabel_clusters_by_recency(rfm_b2c.assign(Cluster=raw_cluster)); sf = silhouette_score(X, raw_cluster)
    rfm_b2c["R_norm_inv"] = -rfm_b2c["R_norm"]; rfm_b2c["R_score"] = pd.qcut(rfm_b2c["R_norm_inv"].rank(method="first"), q=5, labels=RFM_SCORE_LABELS).astype(int)
    rfm_b2c["F_score"] = pd.qcut(rfm_b2c["F_norm"].rank(method="first"), q=5, labels=RFM_SCORE_LABELS).astype(int); rfm_b2c["M_score"] = pd.qcut(rfm_b2c["M_norm"].rank(method="first"), q=5, labels=RFM_SCORE_LABELS).astype(int); rfm_b2c["FM_score"] = (rfm_b2c["F_score"] + rfm_b2c["M_score"]) / 2
    cl = _build_cluster_profile(rfm_b2c); rfm_b2c["Segment"] = rfm_b2c["Cluster"].map(cl.set_index("Cluster")["Segment"])
    out = {"rfm_b2c":rfm_b2c,"cl":cl,"d_raw":d_raw,"d_bc":bc,"d_norm":d_norm,"lams":lams,"X":X,"ine":ine,"sil":sil,"bk":bk,"sf":sf}
    if include_ranges:
        df_q = rfm_b2c.copy(); df_q["R_score_base"] = 1 - df_q["R_norm"]
        out.update({"R_ranges":get_quintile_ranges_ranked(df_q["R_score_base"]),"F_ranges":get_quintile_ranges_ranked(df_q["F_norm"]),"M_ranges":get_quintile_ranges_ranked(df_q["M_norm"])})
    return out
@st.cache_data(show_spinner="📂 Memuat data…")
def _load_bytes():
    if _MODULES_OK: return load_gabungan_bytes()
    p = BASE_DIR / "Data" / "Data hasil pengolahan" / "Transaksi_Gabungan.xlsx"
    return p.read_bytes() if p.exists() else None
@st.cache_data(show_spinner="⏳ Pipeline utama…")
def pipeline(data_bytes:bytes)->dict:
    df=pd.read_excel(io.BytesIO(data_bytes)); df["Waktu Pemesanan"]=pd.to_datetime(df["Waktu Pemesanan"],errors="coerce"); df=df.dropna(subset=["Waktu Pemesanan"]); SNAP=df["Waktu Pemesanan"].max()+pd.Timedelta(days=1)
    rfm=build_rfm(df, SNAP, {"Waktu_Terakhir":("Waktu Pemesanan","max"),"No_Pesanan_Terakhir":("No. Pesanan","last"),"Kanal":("Kanal","last")}); rfm["B2B_flag"]=rfm["Monetary"]>=B2B_THRESHOLD; rfm_b2b=rfm[rfm["B2B_flag"]].copy()
    prep=prepare_b2c_clustering(rfm[~rfm["B2B_flag"]].copy(), range(2,11)); rfm_b2c, cl = prep["rfm_b2c"], prep["cl"]; rfm_b2b["Cluster"]=-1; rfm_b2b["Segment"]="B2B"; rfm_b2b[["R_norm","F_norm","M_norm"]]=None
    return {"df":df,"rfm":rfm,"rfm_b2c":rfm_b2c,"rfm_b2b":rfm_b2b,"rfm_fin":pd.concat([rfm_b2c,rfm_b2b],ignore_index=True),"lams":prep["lams"],"X":prep["X"],"KR":range(2,11),"ine":prep["ine"],"sil":prep["sil"],"sf":prep["sf"],"bk":prep["bk"],"cl":cl,"SNAP":SNAP,"d_raw":prep["d_raw"].copy(),"d_bc":prep["d_bc"].copy(),"d_norm":prep["d_norm"].copy()}
@st.cache_data(show_spinner="⏳ Pipeline B2C (rentang data)…")
def pipeline_b2c_filtered(data_bytes:bytes, date_from, date_to)->dict:
    df_all=pd.read_excel(io.BytesIO(data_bytes)); df_all["Waktu Pemesanan"]=pd.to_datetime(df_all["Waktu Pemesanan"],errors="coerce"); df_all=df_all.dropna(subset=["Waktu Pemesanan"])
    df=df_all[(df_all["Waktu Pemesanan"].dt.date>=date_from)&(df_all["Waktu Pemesanan"].dt.date<=date_to)].copy()
    if df.empty: return {}
    SNAP=df["Waktu Pemesanan"].max()+pd.Timedelta(days=1); rfm=build_rfm(df, SNAP, {"Waktu_Terakhir":("Waktu Pemesanan","max"),"No_Pesanan_Terakhir":("No. Pesanan","last"),"Kanal":("Kanal","last")}); rfm["B2B_flag"]=rfm["Monetary"]>=B2B_THRESHOLD; rfm_b2c=rfm[~rfm["B2B_flag"]].copy()
    if len(rfm_b2c)<10: return {}
    KR=range(2,min(11,len(rfm_b2c)))
    if len(KR)==0: return {}
    prep=prepare_b2c_clustering(rfm_b2c, KR, safe_boxcox=True, include_ranges=True)
    return {"rfm_b2c":prep["rfm_b2c"],"cl":prep["cl"],"SNAP":SNAP,"sf":prep["sf"],"bk":prep["bk"],"d_raw":prep["d_raw"],"d_bc":prep["d_bc"],"d_norm":prep["d_norm"],"lams":prep["lams"],"X":prep["X"],"KR":KR,"ine":prep["ine"],"sil":prep["sil"],"R_ranges":prep["R_ranges"],"F_ranges":prep["F_ranges"],"M_ranges":prep["M_ranges"]}
def dyn_rfm(df_p,snap):
    if df_p is None or df_p.empty: return pd.DataFrame()
    try:
        lt=(df_p.sort_values("Waktu Pemesanan").drop_duplicates("Username (Pembeli)",keep="last")[["Username (Pembeli)","User ID","No. Pesanan","Kanal","Waktu Pemesanan"]].rename(columns={"Waktu Pemesanan":"Terbaru"}))
        rfm=df_p.groupby("Username (Pembeli)").agg(Recency=("Waktu Pemesanan",lambda x:(snap-x.max()).days), Frequency=("No. Pesanan","nunique"), Monetary=("Total Harga Produk","sum")).reset_index().merge(lt,on="Username (Pembeli)",how="left")
        rfm=rfm[rfm["Monetary"]>0].copy()
        if len(rfm)<10: return pd.DataFrame()
        rfm["R_score"]=pd.qcut(rfm["Recency"].rank(method="first"),5,labels=[5,4,3,2,1]).astype(int); rfm["F_score"]=pd.qcut(rfm["Frequency"].rank(method="first"),5,labels=[1,2,3,4,5]).astype(int); rfm["M_score"]=pd.qcut(rfm["Monetary"].rank(method="first"),5,labels=[1,2,3,4,5]).astype(int); rfm["FM_score"]=(rfm["F_score"]+rfm["M_score"])/2; rfm["Segment"]=rfm.apply(lambda x:_seg(x["R_score"],x["FM_score"]),axis=1)
        return rfm
    except Exception: return pd.DataFrame()
for k_,v_ in [("auth",False),("uname",""),("urole",""),("log",[]),("strat_log",[]),("ck",{})]:
    if k_ not in st.session_state: st.session_state[k_]=v_
if not st.session_state.auth:
    _,cc,_=st.columns([1,1.1,1])
    with cc:
        if os.path.exists(LOGO_PATH): st.image(LOGO_PATH,width=160)
        st.markdown("<div style='text-align:center;padding:1rem 0 1.5rem'><div style='font-size:1.8rem;font-weight:800'>Charmigos CRM</div><div style='color:gray;font-size:.9rem'>Sistem Informasi Pengelolaan Pelanggan</div></div>",unsafe_allow_html=True)
        with st.form("lg"):
            un=st.text_input("Username",placeholder="manager / staff"); pw=st.text_input("Password",type="password")
            if st.form_submit_button("🔐 Masuk",use_container_width=True):
                u=_check(un,pw)
                if u: st.session_state.auth=True; st.session_state.uname=u["name"]; st.session_state.urole=u["role"]; add_log("Login"); st.rerun()
                else: st.error("Username atau password salah.")
        with st.expander("Hint Login"): st.caption("Demo — manager: `charmigos123`  |  staff: `staff123`")
    st.stop()
with st.sidebar:
    if os.path.exists(LOGO_PATH): st.image(LOGO_PATH,width=130)
    else: st.markdown("## 🛍️ Charmigos CRM")
    st.markdown(f"**{st.session_state.uname}** · {st.session_state.urole}")
    data_bytes=_load_bytes(); st.markdown("---")
    if data_bytes: st.success("✅ Data aktif")
    else: st.error("❌ Data belum tersedia")
    st.markdown("---"); page=st.radio("Navigasi",["🏠 Home","🎯 Segmentasi Pelanggan","📋 Strategi Pengelolaan","📥 Data & Input"]); st.markdown("---")
    if st.button("🚪 Keluar"): add_log("Logout"); st.session_state.auth=False; st.rerun()
    st.caption("© Charmigos CRM 2026")
def need_data():
    st.info("📂 Data belum tersedia. Upload di **Data & Input**.",icon="ℹ️"); st.stop()
if page=="🏠 Home":
    st.title("🏠 Home")
    if not data_bytes: need_data()
    df_all=pd.read_excel(io.BytesIO(data_bytes))
    df_all["Waktu Pemesanan"]=pd.to_datetime(df_all["Waktu Pemesanan"],errors="coerce")
    df_all=df_all.dropna(subset=["Waktu Pemesanan"])
    mn,mx=df_all["Waktu Pemesanan"].min().date(),df_all["Waktu Pemesanan"].max().date()
    fc1,fc2=st.columns(2)
    d_from=fc1.date_input("📅 Dari",value=mn,min_value=mn,max_value=mx)
    d_to  =fc2.date_input("📅 Sampai",value=mx,min_value=mn,max_value=mx)
    dff=df_all[(df_all["Waktu Pemesanan"].dt.date>=d_from)&(df_all["Waktu Pemesanan"].dt.date<=d_to)].copy()
    st.markdown("---")

    rfm_f=(dff.groupby("User ID").agg(Monetary_Total=("Total Harga Produk","sum")).reset_index()); n_b2c=int((rfm_f["Monetary_Total"]<B2B_THRESHOLD).sum()); n_b2b=int((rfm_f["Monetary_Total"]>=B2B_THRESHOLD).sum()); n_uniq=int(dff["User ID"].nunique())
    _b2c_users_home = set(rfm_f.loc[rfm_f["Monetary_Total"]<B2B_THRESHOLD,"User ID"]); dff_b2c = dff[dff["User ID"].isin(_b2c_users_home)].copy()
    _df_all_to = df_all[df_all["Waktu Pemesanan"].dt.date <= d_to]; _tx_kumulatif = _df_all_to.groupby("User ID")["No. Pesanan"].nunique(); _buyers_rentang = set(dff["User ID"].unique()); n_repeat = int(sum(1 for uid in _buyers_rentang if _tx_kumulatif.get(uid,0) >= 2)); n_b2c_uniq = len(_b2c_users_home)
    k1,k2,k3,k4=st.columns(4)
    k1.metric("📦 Jumlah Transaksi",f"{len(dff):,}")
    k2.metric("👥 Jumlah Pelanggan Keseluruhan",f"{n_uniq:,}")
    k3.metric("💰 Revenue",f"Rp {dff['Total Harga Produk'].sum()/1e6:.1f}jt")
    k4.metric("🛒 Rata-rata Pembelian",f"Rp {dff['Total Harga Produk'].mean():,.0f}" if len(dff) else "Rp 0")
    k5,k6,k7=st.columns(3)
    k5.metric("🏪 Pelanggan B2C",f"{n_b2c:,}",f"{n_b2c/n_uniq:.1%}" if n_uniq else "")
    k6.metric("🏢 Pelanggan B2B",f"{n_b2b:,}",f"{n_b2b/n_uniq:.1%}" if n_uniq else "")
    _rpr_pct_home = n_repeat/n_uniq*100 if n_uniq else 0
    k7.metric("🔄 Repeat Buyer",f"{n_repeat:,}",f"RPR Keseluruhan Periode: {_rpr_pct_home:.2f}%" if n_uniq else "")
    st.caption(f"Data: **{d_from}** s/d **{d_to}**  |  RPR = {n_repeat:,} / {n_uniq:,} = **{_rpr_pct_home:.2f}%**  (Pembeli Berulang = semua pelanggan B2C & B2B dengan ≥2 transaksi kumulatif s/d {d_to})")
    st.markdown("---")
    st.subheader("📊 Revenue Bulanan")
    dff2=dff.copy(); dff2["Bulan"]=dff2["Waktu Pemesanan"].dt.to_period("M").astype(str); rev_k=dff2.groupby(["Bulan","Kanal"])["Total Harga Produk"].sum().reset_index(); rev_k.columns=["Bulan","Kanal","Revenue"]
    kanal_src=dff.copy(); kanal_src["Kanal"]=kanal_src["Kanal"].fillna("Tidak Diketahui"); kanal_src["Total Harga Produk"]=pd.to_numeric(kanal_src["Total Harga Produk"], errors="coerce").fillna(0)
    kanal_tx=kanal_src.groupby("Kanal", dropna=False).agg(Jumlah_Transaksi=("No. Pesanan","nunique"), Revenue=("Total Harga Produk","sum")).reset_index()
    kanal_tx["Jumlah_Transaksi"]=pd.to_numeric(kanal_tx["Jumlah_Transaksi"], errors="coerce").fillna(0).astype(int); kanal_tx["Revenue"]=pd.to_numeric(kanal_tx["Revenue"], errors="coerce").fillna(0.0); kanal_tx=kanal_tx[kanal_tx["Jumlah_Transaksi"]>0].copy()
    total_tx_periode=int(kanal_tx["Jumlah_Transaksi"].sum()) if not kanal_tx.empty else 0; kanal_tx["Total_Transaksi_Periode"]=total_tx_periode
    kanal_tx["Hover_Text"]=kanal_tx.apply(lambda row: f"<b>{row['Kanal']}</b><br>Jumlah Transaksi Kanal: {int(row['Jumlah_Transaksi']):,}<br>Total Transaksi Periode: {int(row['Total_Transaksi_Periode']):,}<br>Revenue Kanal: Rp {float(row['Revenue']):,.0f}", axis=1)
    rv1,rv2=st.columns([1.6,0.9])
    with rv1:
        fig_rev=px.bar(rev_k,x="Bulan",y="Revenue",color="Kanal",barmode="stack", color_discrete_map=KANAL_COLOR,labels={"Revenue":"Revenue (Rp)"})
        fig_rev.update_layout(height=340,legend=dict(orientation="h",yanchor="bottom",y=1.02,x=1,xanchor="right"), margin=dict(t=40,b=50,l=0,r=0),xaxis_tickangle=-30,yaxis_tickformat=",.0f")
        fig_rev.update_traces(hovertemplate="<b>%{x}</b><br>%{fullData.name}: Rp %{y:,.0f}<extra></extra>")
        st.plotly_chart(fig_rev,use_container_width=True)
    with rv2:
        if kanal_tx.empty:
            st.info("Belum ada data transaksi per kanal pada rentang ini.")
        else:
            fig_kanal=go.Figure(data=[go.Pie(labels=kanal_tx["Kanal"], values=kanal_tx["Jumlah_Transaksi"], hovertext=kanal_tx["Hover_Text"], textinfo="percent+label", textposition="inside", marker=dict(colors=[KANAL_COLOR.get(k, "#94a3b8") for k in kanal_tx["Kanal"]]), hovertemplate="%{hovertext}<br>Komposisi: %{percent}<extra></extra>")])
            fig_kanal.update_layout(title="Distribusi Transaksi per Kanal", height=340, margin=dict(t=60,b=20,l=0,r=0), showlegend=False)
            st.plotly_chart(fig_kanal,use_container_width=True)
    st.markdown("---")

    st.subheader("🔄 Repeat Purchase Rate per Bulan")
    df_rpr_base = dff.copy()
    df_rpr_base["Bulan"] = df_rpr_base["Waktu Pemesanan"].dt.to_period("M")
    df_all_rpr  = df_all.copy()
    df_all_rpr["Bulan"] = df_all_rpr["Waktu Pemesanan"].dt.to_period("M")
    months_sorted = sorted(df_rpr_base["Bulan"].unique())
    rpr_rows = []
    for m in months_sorted:
        buyers_m = df_rpr_base[df_rpr_base["Bulan"] == m]["User ID"].unique(); total_buyers = len(buyers_m)
        if total_buyers == 0: continue
        _df_kum_m = df_all_rpr[df_all_rpr["Bulan"] <= m]; tx_kum = _df_kum_m.groupby("User ID")["No. Pesanan"].nunique(); _monetary_kum_m = _df_kum_m.groupby("User ID")["Total Harga Produk"].sum()
        n_repeat_m = int(sum(1 for uid in buyers_m if tx_kum.get(uid, 0) >= 2)); n_new_m = total_buyers - n_repeat_m; rpr_pct = n_repeat_m / total_buyers * 100
        n_repeat_b2b = int(sum(1 for uid in buyers_m if tx_kum.get(uid, 0) >= 2 and _monetary_kum_m.get(uid, 0) >= B2B_THRESHOLD)); n_repeat_b2c = n_repeat_m - n_repeat_b2b
        rpr_rows.append({"Bulan": str(m), "Pembeli Baru": n_new_m, "Pembeli Berulang": n_repeat_m, "  ↳ B2C Berulang": n_repeat_b2c, "  ↳ B2B Berulang": n_repeat_b2b, "Total Pembeli": total_buyers, "Repeat Purchase Rate (%)": round(rpr_pct, 2)})
    df_rpr_tbl = pd.DataFrame(rpr_rows)
    fig_rpr = go.Figure()
    fig_rpr.add_trace(go.Bar(x=df_rpr_tbl["Bulan"], y=df_rpr_tbl["Repeat Purchase Rate (%)"], marker_color="#3b82f6", text=df_rpr_tbl["Repeat Purchase Rate (%)"].apply(lambda v: f"{v:.2f}%"), textposition="outside", hovertemplate="<b>%{x}</b><br>RPR: %{y:.2f}%<br>Berulang: %{customdata[0]:,} | Total: %{customdata[1]:,}<extra></extra>", customdata=df_rpr_tbl[["Pembeli Berulang","Total Pembeli"]].values))
    _rpr_max = df_rpr_tbl["Repeat Purchase Rate (%)"].max() if not df_rpr_tbl.empty else 5
    fig_rpr.update_layout(height=360, xaxis=dict(title="Bulan", tickangle=-30, gridcolor="#eee"), yaxis=dict(title="RPR (%)", gridcolor="#eee", range=[0, max(_rpr_max*1.25, 5)]), margin=dict(t=30, b=50, l=0, r=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    st.plotly_chart(fig_rpr, use_container_width=True)

    with st.expander("📋 Lihat Detail — Tabel RPR per Bulan", expanded=False):
        df_rpr_disp = df_rpr_tbl.copy()
        df_rpr_disp["Repeat Purchase Rate (%)"] = df_rpr_disp["Repeat Purchase Rate (%)"].apply(lambda v: f"{v:.2f}%")
        st.dataframe(df_rpr_disp, use_container_width=True, hide_index=True)
        maybe_sync_download(
            st.download_button("📥 Download Tabel RPR", data=to_xl(df_rpr_tbl, "Repeat_Purchase_Rate"), file_name="Repeat_Purchase_Rate.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            df_rpr_tbl,
            "Repeat_Purchase_Rate.xlsx",
        )
    st.caption("💡 **Pembeli Berulang** = pelanggan yang bertransaksi bulan ini dengan ≥2 transaksi kumulatif (seluruh data historis) s/d bulan tersebut. Definisi sama dengan kartu RPR di atas.")
    st.markdown("---")

    st.subheader("💰 Kontribusi Revenue dari Pelanggan Berulang")
    _freq_in_range = dff.groupby("User ID")["No. Pesanan"].nunique(); _buyers_1x = set(_freq_in_range[_freq_in_range == 1].index); _buyers_2x = set(_freq_in_range[_freq_in_range >= 2].index)
    dff_rev = dff.copy(); dff_rev["_freq_group"] = dff_rev["User ID"].map(lambda uid: "2x+" if uid in _buyers_2x else "1x")
    _rev_2x = dff_rev.loc[dff_rev["_freq_group"] == "2x+", "Total Harga Produk"].sum(); _rev_1x = dff_rev.loc[dff_rev["_freq_group"] == "1x", "Total Harga Produk"].sum(); _rev_total = _rev_2x + _rev_1x
    _n_2x, _n_1x, _n_total = len(_buyers_2x), len(_buyers_1x), len(_buyers_2x) + len(_buyers_1x)
    _avg_2x = _rev_2x / _n_2x if _n_2x else 0; _avg_1x = _rev_1x / _n_1x if _n_1x else 0; _avg_total = _rev_total / _n_total if _n_total else 0
    _pct_2x = _rev_2x / _rev_total * 100 if _rev_total else 0; _pct_1x = _rev_1x / _rev_total * 100 if _rev_total else 0; _mix_2x = _n_2x / _n_total * 100 if _n_total else 0; _mix_1x = _n_1x / _n_total * 100 if _n_total else 0
    _rev_contrib_df = pd.DataFrame([
        {"Jenis Pelanggan": "🔄 Pelanggan Berulang", "Jumlah Pelanggan": _n_2x, "Komposisi Pelanggan (%)": f"{_mix_2x:.1f}%", "Total Nilai Transaksi": f"Rp {_rev_2x:,.0f}", "Rata-rata per Pelanggan": f"Rp {_avg_2x:,.0f}", "Kontribusi (%)": f"{_pct_2x:.1f}%"},
        {"Jenis Pelanggan": "1️⃣ Pelanggan Sekali Beli", "Jumlah Pelanggan": _n_1x, "Komposisi Pelanggan (%)": f"{_mix_1x:.1f}%", "Total Nilai Transaksi": f"Rp {_rev_1x:,.0f}", "Rata-rata per Pelanggan": f"Rp {_avg_1x:,.0f}", "Kontribusi (%)": f"{_pct_1x:.1f}%"},
        {"Jenis Pelanggan": "📊 Total", "Jumlah Pelanggan": _n_total, "Komposisi Pelanggan (%)": "100%", "Total Nilai Transaksi": f"Rp {_rev_total:,.0f}", "Rata-rata per Pelanggan": f"Rp {_avg_total:,.0f}", "Kontribusi (%)": "100%"},
    ])
    _rc1, _rc2 = st.columns([1.1, 0.9])
    with _rc1:
        st.dataframe(_rev_contrib_df, use_container_width=True, hide_index=True)
        st.caption("💡 **Pelanggan Berulang** = pelanggan dengan ≥2 transaksi dalam rentang. **Pelanggan Sekali Beli** = pelanggan dengan hanya 1 transaksi dalam rentang.")
    with _rc2:
        _bar_df = pd.DataFrame({
            "Kategori": ["Pelanggan Berulang", "Pelanggan Sekali Beli"],
            "Revenue": [_rev_2x, _rev_1x],
            "Jumlah": [_n_2x, _n_1x],
            "Rata-rata": [_avg_2x, _avg_1x],
            "Pct": [_pct_2x, _pct_1x],
        })
        _fig_bar = go.Figure()
        _fig_bar.add_trace(go.Bar(x=_bar_df["Revenue"], y=_bar_df["Kategori"], orientation="h", marker_color=["#3b82f6", "#f59e0b"], text=_bar_df["Pct"].apply(lambda v: f"{v:.1f}%"), textposition="inside", insidetextanchor="middle", textfont=dict(color="white", size=13, family="Arial"), customdata=_bar_df[["Jumlah","Rata-rata","Pct"]].values, hovertemplate="<b>%{y}</b><br>Total Revenue: Rp %{x:,.0f}<br>Jumlah Pelanggan: %{customdata[0]:,}<br>Rata-rata per Pelanggan: Rp %{customdata[1]:,.0f}<br>Kontribusi: %{customdata[2]:.1f}%<extra></extra>"))
        _fig_bar.update_layout(height=180, xaxis=dict(title="Total Revenue (Rp)", tickformat=",.0f", gridcolor="#eee"), yaxis=dict(autorange="reversed"), margin=dict(t=10, b=40, l=0, r=10), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(_fig_bar, use_container_width=True)

    st.markdown("---")

    st.subheader("🎯 Profil Klaster B2C")
    with st.spinner("Menghitung klaster untuk rentang tanggal ini..."): p_home=pipeline_b2c_filtered(data_bytes, d_from, d_to)
    if p_home:
        cl_h=p_home["cl"]
        cl_d=cl_h[["Cluster","Count","Recency","Frequency","Monetary","R_score","FM_score","Segment","Strategi"]].rename(columns={"Count":"Pelanggan","Recency":"Avg Recency","Frequency":"Avg Freq","Monetary":"Avg Monetary"})
        cl_d["Avg Recency"]=cl_d["Avg Recency"].round(1)
        cl_d["Avg Freq"]=cl_d["Avg Freq"].round(2)
        cl_d["Avg Monetary"]=cl_d["Avg Monetary"].apply(lambda x:f"Rp {x:,.0f}")
        st.dataframe(cl_d,use_container_width=True,hide_index=True)
    else:
        st.info("Data B2C pada rentang ini terlalu sedikit untuk clustering.")
    st.markdown("---")

    st.subheader("🏢 Profil Pelanggan B2B")
    rfm_b2b_f=(dff.groupby("User ID").agg(Username=("Username (Pembeli)","first"), Monetary=("Total Harga Produk","sum"), Frequency=("No. Pesanan","nunique"), Recency=("Waktu Pemesanan",lambda x:(pd.Timestamp(d_to)-x.max()).days), Kanal=("Kanal","last")).reset_index())
    rfm_b2b_f=rfm_b2b_f[rfm_b2b_f["Monetary"]>=B2B_THRESHOLD]
    if not rfm_b2b_f.empty:
        b1,b2,b3,b4,b5=st.columns(5)
        b1.metric("Jumlah Pelanggan B2B",f"{len(rfm_b2b_f):,}")
        b2.metric("Total Revenue B2B",f"Rp {rfm_b2b_f['Monetary'].sum():,.0f}")
        b3.metric("Avg Monetary",f"Rp {rfm_b2b_f['Monetary'].mean():,.0f}")
        b4.metric("Max Monetary",f"Rp {rfm_b2b_f['Monetary'].max():,.0f}")
        b5.metric("Min Monetary",f"Rp {rfm_b2b_f['Monetary'].min():,.0f}")
        b2b_by_kanal=rfm_b2b_f.groupby("Kanal").agg(Pelanggan=("User ID","count"), Total_Revenue=("Monetary","sum"), Avg_Monetary=("Monetary","mean"), Max_Monetary=("Monetary","max"), Min_Monetary=("Monetary","min")).round(0).reset_index()
        st.dataframe(b2b_by_kanal,use_container_width=True,hide_index=True)
    else:
        st.info("Tidak ada pelanggan B2B pada rentang ini.")
    st.markdown("---")

    st.subheader("📋 Log Aktivitas")
    if _MODULES_OK:
        st.caption(f"📁 Data tersimpan di: `{PROCESSED_DIR}`")
    if st.session_state.log:
        ldf=pd.DataFrame(st.session_state.log); ldf["Waktu"]=pd.to_datetime(ldf["Waktu"])
        lf1,lf2,lf3=st.columns(3)
        l_from=lf1.date_input("Log dari:",value=ldf["Waktu"].min().date(),key="lf")
        l_to  =lf2.date_input("Log sampai:",value=ldf["Waktu"].max().date(),key="lt")
        l_usr =lf3.selectbox("Pengguna:",["Semua"]+sorted(ldf["Pengguna"].unique()),key="lu")
        lf=ldf[(ldf["Waktu"].dt.date>=l_from)&(ldf["Waktu"].dt.date<=l_to)]
        if l_usr!="Semua": lf=lf[lf["Pengguna"]==l_usr]
        lf["Waktu"]=lf["Waktu"].dt.strftime("%Y-%m-%d %H:%M:%S")
        lf_show=lf.iloc[::-1].reset_index(drop=True)
        st.dataframe(lf_show,use_container_width=True,height=200)
        maybe_sync_download(
            st.download_button("📥 Export Log",data=to_xl(lf_show,"Log_Aktivitas"), file_name="Log_Aktivitas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            lf_show,
            "Log_Aktivitas.xlsx",
        )
    else: st.info("Belum ada aktivitas.")

elif page=="🎯 Segmentasi Pelanggan":
    st.title("🎯 Segmentasi Pelanggan")
    if not data_bytes: need_data()
    add_log("Segmentasi")

    df_seg_all = pd.read_excel(io.BytesIO(data_bytes))
    df_seg_all["Waktu Pemesanan"] = pd.to_datetime(df_seg_all["Waktu Pemesanan"], errors="coerce")
    df_seg_all = df_seg_all.dropna(subset=["Waktu Pemesanan"])
    _seg_mn = df_seg_all["Waktu Pemesanan"].min().date()
    _seg_mx = df_seg_all["Waktu Pemesanan"].max().date()

    sc1, sc2 = st.columns(2)
    _seg_d1 = sc1.date_input("📅 Dari", value=_seg_mn, min_value=_seg_mn, max_value=_seg_mx, key="seg_d1")
    _seg_d2 = sc2.date_input("📅 Sampai", value=_seg_mx, min_value=_seg_mn, max_value=_seg_mx, key="seg_d2")
    if _seg_d1 > _seg_d2:
        st.error("Tanggal Dari harus sebelum Sampai."); st.stop()

    df_seg_f = df_seg_all[(df_seg_all["Waktu Pemesanan"].dt.date >= _seg_d1) & (df_seg_all["Waktu Pemesanan"].dt.date <= _seg_d2)].copy()
    if df_seg_f.empty:
        st.warning("Tidak ada data pada rentang ini."); st.stop()

    st.markdown("---")

    with st.spinner("⏳ Membangun RFM & klaster..."):
        _p_seg = pipeline_b2c_filtered(data_bytes, _seg_d1, _seg_d2)
    _SNAP_seg = df_seg_f["Waktu Pemesanan"].max() + pd.Timedelta(days=1)
    rfm_seg = build_rfm(df_seg_f, _SNAP_seg, {"Waktu_Terakhir": ("Waktu Pemesanan", "max"), "No_Pesanan_Terakhir": ("No. Pesanan", "last"), "Kanal": ("Kanal", "last")})
    rfm_seg["B2B_flag"] = rfm_seg["Monetary"] >= B2B_THRESHOLD
    rfm_b2b_seg = rfm_seg[rfm_seg["B2B_flag"]].copy()
    rfm_b2c_seg = rfm_seg[~rfm_seg["B2B_flag"]].copy()

    if _p_seg:
        _seg_map = _p_seg["rfm_b2c"].set_index("User ID")["Segment"].to_dict()
        _clu_map = _p_seg["rfm_b2c"].set_index("User ID")["Cluster"].to_dict()
        rfm_seg["Segment"] = rfm_seg["User ID"].map(_seg_map)
        rfm_seg["Cluster"] = rfm_seg["User ID"].map(_clu_map)
        rfm_seg.loc[rfm_seg["B2B_flag"], "Segment"] = "B2B"
        rfm_b2c_seg["Segment"] = rfm_b2c_seg["User ID"].map(_seg_map)
        rfm_b2c_seg["Cluster"] = rfm_b2c_seg["User ID"].map(_clu_map)

    st.header("Pembentukan Variabel RFM")
    st.caption(f"Snapshot: `{_SNAP_seg.date()}` | Rentang: **{_seg_d1}** s/d **{_seg_d2}** | Total: **{len(rfm_seg):,}** pelanggan")
    o1, o2, o3 = st.columns(3)
    o1.metric("👥 Total Pelanggan", f"{len(rfm_seg):,}")
    o2.metric("📦 Total Transaksi", f"{len(df_seg_f):,}")
    o3.metric("📅 Rentang Data", f"{_seg_d1} – {_seg_d2}")
    with st.expander(f"📋 Tabel RFM — {len(rfm_seg):,} pelanggan", expanded=False):
        rfm_show = rfm_seg.reset_index(drop=True); rfm_show.index += 1
        st.dataframe(rfm_show[[c for c in ["Username","Recency","Frequency","Monetary","No_Pesanan_Terakhir","Waktu_Terakhir","Kanal","Segment"] if c in rfm_show.columns]], use_container_width=True, height=280)
        maybe_sync_download(
            st.download_button("📥 Download RFM", data=to_xl(rfm_show.reset_index(), "RFM", ["User ID"]), file_name="RFM_Transaksi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            rfm_show.reset_index(),
            "RFM_Transaksi.xlsx",
        )
    st.markdown("---")
    st.header(" Segmen B2B")
    st.caption(f"Batasan B2B: Monetary ≥ **Rp {B2B_THRESHOLD:,.0f}** — tidak ikut K-Means")
    bb1, bb2 = st.columns(2)
    bb1.metric("🏢 B2B", f"{len(rfm_b2b_seg):,}", f"{len(rfm_b2b_seg)/len(rfm_seg):.1%}" if len(rfm_seg) else "")
    bb2.metric("🏪 B2C", f"{len(rfm_b2c_seg):,}", f"{len(rfm_b2c_seg)/len(rfm_seg):.1%}" if len(rfm_seg) else "")
    with st.expander(f"📋 Data B2B — {len(rfm_b2b_seg):,} pelanggan", expanded=False):
        b2b_show = rfm_b2b_seg.reset_index(drop=True); b2b_show.index += 1
        st.dataframe(b2b_show[[c for c in ["Username","Recency","Frequency","Monetary","No_Pesanan_Terakhir","Waktu_Terakhir","Kanal"] if c in b2b_show.columns]], use_container_width=True, height=240)
        maybe_sync_download(
            st.download_button("📥 Download Segmen B2B", data=to_xl(b2b_show.reset_index(), "Data_B2B_Export", ["User ID"]), file_name="Segmen_B2B.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            b2b_show.reset_index(),
            "Segmen_B2B.xlsx",
        )
    st.markdown("---")
    st.header(" Segmen B2C")
    if not _p_seg:
        st.warning("Data B2C terlalu sedikit untuk dianalisis pada rentang ini."); st.stop()
    rfm_b2c_f = _p_seg["rfm_b2c"]; cl_f = _p_seg["cl"]
    fb = _p_seg  # alias agar kode tab di bawah tetap kompatibel
    st.info(f"**{len(rfm_b2c_f):,}** pelanggan B2C | K optimal: **{_p_seg['bk']}** klaster | Silhouette: **{_p_seg['sf']:.4f}**", icon="📊")
    tab_norm, tab_elbow, tab_vis = st.tabs(["📊 Normalisasi", "📐 Penentuan Klaster", "📈 Visualisasi Klaster"])
    with tab_norm:
        st.subheader("Box-Cox → MinMax (B2C)")
        lams_f = fb["lams"]
        lc1, lc2 = st.columns(2)
        if lams_f.get("Recency"):  lc1.metric("λ Box-Cox Recency",  f"{lams_f['Recency']:.4f}")
        if lams_f.get("Monetary"): lc2.metric("λ Box-Cox Monetary", f"{lams_f['Monetary']:.4f}")
        feats  = ["Recency", "Frequency", "Monetary"]
        stages = [("Raw", fb["d_raw"]), ("Box-Cox", fb["d_bc"]), ("MinMax (0–1)", fb["d_norm"])]
        fig_hist = make_subplots(rows=3, cols=3, subplot_titles=[f"{feat} ({s})" for s in [x[0] for x in stages] for feat in feats])
        colors = ["#4C72B0", "#55A868", "#C44E52"]
        for ri, (stage_name, d) in enumerate(stages):
            for ci, feat in enumerate(feats):
                v = d[feat].dropna()
                fig_hist.add_trace(go.Histogram(x=v, nbinsx=40, name=f"{stage_name} {feat}", marker_color=colors[ci], opacity=0.8, showlegend=False, hovertemplate=f"<b>{stage_name} {feat}</b><br>Range: %{{x}}<br>Count: %{{y}}<extra></extra>"), row=ri+1, col=ci+1)
        fig_hist.update_layout(height=600, title_text="Transformasi RFM — hover untuk detail", margin=dict(t=80, b=20))
        st.plotly_chart(fig_hist, use_container_width=True)
        with st.expander("📋 Tabel Data Ternormalisasi (preview)", expanded=False):
            n_show = rfm_b2c_f[["User ID","Username","Recency","Frequency","Monetary","R_norm","F_norm","M_norm"]].copy()
            n_show = n_show.reset_index(drop=True); n_show.index += 1
            st.dataframe(n_show.round(4), use_container_width=True, height=240)
            maybe_sync_download(
                st.download_button("📥 Download", data=to_xl(n_show.reset_index(), "RFM_Normalized", ["User ID"]), file_name="RFM_Normalized.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                n_show.reset_index(),
                "RFM_Normalized.xlsx",
            )
        with st.expander("📐 Rentang Kuintil (dari distribusi data aktual)", expanded=False):
            fmt = lambda x: f"{x:.4f}"
            q_table = []
            for i in range(5):
                q_table.append({"Skor": i+1, "R_norm (dibalik)": f"{fmt(fb['R_ranges'][4-i][0])} – {fmt(fb['R_ranges'][4-i][1])}", "F_norm": f"{fmt(fb['F_ranges'][i][0])} – {fmt(fb['F_ranges'][i][1])}", "M_norm": f"{fmt(fb['M_ranges'][i][0])} – {fmt(fb['M_ranges'][i][1])}"})
            st.caption("Rentang kuintil dihitung ulang sesuai distribusi data B2C pada rentang yang dipilih.")
            st.dataframe(pd.DataFrame(q_table).set_index("Skor"), use_container_width=True)

    with tab_elbow:
        st.subheader("Elbow Method — K Optimal")
        KR_f, ine_f, sil_f = list(fb["KR"]), fb["ine"], fb["sil"]
        fig_e = go.Figure()
        fig_e.add_trace(go.Scatter(x=KR_f, y=ine_f, mode="lines+markers", name="Inertia", line=dict(color="#4C72B0", width=2), yaxis="y1", hovertemplate="k=%{x}<br>Inertia: %{y:,.1f}<extra></extra>"))
        fig_e.add_trace(go.Scatter(x=KR_f, y=sil_f, mode="lines+markers", name="Silhouette", line=dict(color="#C44E52", width=2, dash="dash"), marker=dict(symbol="square"), yaxis="y2", hovertemplate="k=%{x}<br>Silhouette: %{y:.4f}<extra></extra>"))
        fig_e.add_vline(x=fb["bk"], line_dash="dot", line_color="orange", annotation_text=f"k={fb['bk']} (optimal)")
        fig_e.update_layout(height=340, title=f"K optimal = {fb['bk']} (Silhouette terbaik)", xaxis=dict(title="k", tickmode="array", tickvals=KR_f, gridcolor="#eee"), yaxis=dict(title=dict(text="Inertia", font=dict(color="#4C72B0")), tickfont=dict(color="#4C72B0"), gridcolor="#eee"), yaxis2=dict(title=dict(text="Silhouette", font=dict(color="#C44E52")), tickfont=dict(color="#C44E52"), overlaying="y", side="right"), margin=dict(t=50, b=50, l=0, r=60))
        st.plotly_chart(fig_e, use_container_width=True)
        st.metric("📐 Silhouette Final (k optimal)", f"{fb['sf']:.4f}")

    with tab_vis:
        v2d, v3d = st.tabs(["🔵 2D", "🌐 3D"])
        with v2d:
            vc1, vc2 = st.columns(2)
            with vc1:
                fig_s = px.scatter(rfm_b2c_f, x="R_norm", y="M_norm", color=rfm_b2c_f["Cluster"].astype(str), opacity=0.5, title="R_norm vs M_norm", color_discrete_sequence=px.colors.qualitative.Set1)
                fig_s.update_traces(marker=dict(size=4))
                fig_s.update_layout(height=320, margin=dict(t=40, b=20, l=0, r=0))
                st.plotly_chart(fig_s, use_container_width=True)
            with vc2:
                fig_s2 = px.scatter(rfm_b2c_f, x="F_norm", y="M_norm", color=rfm_b2c_f["Cluster"].astype(str), opacity=0.5, title="F_norm vs M_norm", color_discrete_sequence=px.colors.qualitative.Set1)
                fig_s2.update_traces(marker=dict(size=4))
                fig_s2.update_layout(height=320, margin=dict(t=40, b=20, l=0, r=0))
                st.plotly_chart(fig_s2, use_container_width=True)
        with v3d:
            tmp = rfm_b2c_f.copy(); tmp["Cluster"] = tmp["Cluster"].astype(str)
            f3 = px.scatter_3d(tmp, x="R_norm", y="F_norm", z="M_norm", color="Cluster", color_discrete_sequence=px.colors.qualitative.Set1, opacity=0.7)
            f3.update_traces(marker=dict(size=3))
            f3.update_layout(margin=dict(l=0, r=0, b=0, t=30), height=500)
            st.plotly_chart(f3, use_container_width=True)
    st.markdown("---")
    st.subheader("🎯 Profil Klaster & Segmen B2C")
    cl_disp = cl_f[["Cluster","Count","Recency","Frequency","Monetary","R_norm","F_norm","M_norm","R_score","F_score","M_score","FM_score","Segment","Strategi"]].copy()
    cl_disp = cl_disp.rename(columns={"Count":"Pelanggan","Recency":"Avg Recency","Frequency":"Avg Freq","Monetary":"Avg Monetary","R_norm":"Avg R_norm","F_norm":"Avg F_norm","M_norm":"Avg M_norm"})
    cl_disp["Avg Recency"] = cl_disp["Avg Recency"].round(1)
    cl_disp["Avg Freq"]    = cl_disp["Avg Freq"].round(2)
    cl_disp["Avg Monetary"]= cl_disp["Avg Monetary"].apply(lambda x: f"Rp {x:,.0f}")
    for c in ["Avg R_norm","Avg F_norm","Avg M_norm","FM_score"]:
        cl_disp[c] = cl_disp[c].round(4)
    st.dataframe(cl_disp, use_container_width=True, hide_index=True)
    with st.expander("📖 Kategori Segmen RFM"):
        st.dataframe(pd.DataFrame([{"Segmen":s,"R":d["R"],"FM":d["FM"],"Deskripsi":d["desc"],"Strategi":d["strategi"]} for s,d in SEG_DETAIL.items()]), use_container_width=True, hide_index=True)
    st.markdown("---")
    st.subheader("🔍 Eksplorasi Data Pelanggan B2C & B2B")
    ef1, ef2, ef3 = st.columns(3)
    f_tipe  = ef1.selectbox("Tipe:", ["Semua","B2C","B2B"])
    _seg_opts = ["Semua"] + sorted(rfm_seg["Segment"].dropna().unique().tolist())
    f_seg_e = ef2.selectbox("Segmen:", _seg_opts)
    f_sr    = ef3.text_input("Cari Username:")
    exp_filt = rfm_seg.copy()
    if f_tipe == "B2C": exp_filt = exp_filt[~exp_filt["B2B_flag"]]
    elif f_tipe == "B2B": exp_filt = exp_filt[exp_filt["B2B_flag"]]
    if f_seg_e != "Semua": exp_filt = exp_filt[exp_filt["Segment"] == f_seg_e]
    if f_sr: exp_filt = exp_filt[exp_filt["Username"].str.contains(f_sr, case=False, na=False)]
    exp_show = exp_filt.reset_index(drop=True); exp_show.index += 1
    st.write(f"**{len(exp_show):,}** pelanggan:")
    disp_cols = [c for c in ["Username","Recency","Frequency","Monetary","No_Pesanan_Terakhir","Waktu_Terakhir","Cluster","Segment"] if c in exp_show.columns]
    st.dataframe(exp_show[disp_cols], use_container_width=True, height=300)
    st.markdown("---")
    st.subheader("📤 Export Hasil Segmentasi")
    rfm_b2c_e = rfm_b2c_seg; rfm_b2b_e = rfm_b2b_seg; cl_e = cl_f
    sum_b2c = rfm_b2c_e.groupby("Segment").agg(Pelanggan=("User ID","count"), Avg_R=("Recency","mean"), Avg_F=("Frequency","mean"), Avg_M=("Monetary","mean")).round(2).reset_index() if "Segment" in rfm_b2c_e.columns else pd.DataFrame()
    sum_b2b = pd.DataFrame([{"Segment":"B2B","Pelanggan":len(rfm_b2b_e), "Avg_R":rfm_b2b_e["Recency"].mean().round(2) if not rfm_b2b_e.empty else 0, "Avg_F":rfm_b2b_e["Frequency"].mean().round(2) if not rfm_b2b_e.empty else 0, "Avg_M":rfm_b2b_e["Monetary"].mean().round(2) if not rfm_b2b_e.empty else 0}])
    summary_all = pd.concat([sum_b2c, sum_b2b], ignore_index=True)
    strat_df = cl_e[["Cluster","Segment","Count","Recency","Frequency","Monetary","R_score","FM_score","Strategi"]].copy().rename(columns={"Count":"Pelanggan"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for sheet, data in {
            "RFM_Data_Lengkap": rfm_seg[["User ID","Username","Recency","Frequency","Monetary","Cluster","Segment","B2B_flag"]],
            "Data_B2C": rfm_b2c_e[["User ID","Username","Recency","Frequency","Monetary","R_score","F_score","M_score","FM_score","Cluster","Segment"] if all(c in rfm_b2c_e.columns for c in ["R_score","F_score"]) else [c for c in ["User ID","Username","Recency","Frequency","Monetary","Cluster","Segment"] if c in rfm_b2c_e.columns]],
            "Data_B2B": rfm_b2b_e[["User ID","Username","Recency","Frequency","Monetary","Segment"] if "Segment" in rfm_b2b_e.columns else [c for c in ["User ID","Username","Recency","Frequency","Monetary"] if c in rfm_b2b_e.columns]],
            "Profil_Cluster": cl_e[["Cluster","Count","Recency","Frequency","Monetary","R_norm","F_norm","M_norm","R_score","F_score","M_score","FM_score","Segment"]],
            "Strategi_Cluster": strat_df, "Ringkasan_Segmen": summary_all,
        }.items(): data.to_excel(w, sheet_name=sheet, index=False)
    maybe_sync_download(
        st.download_button("📥 Download hasil segmentasi pelanggan", data=buf.getvalue(), file_name="RFM_Clustering_Result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary"),
        rfm_seg[["User ID","Username","Recency","Frequency","Monetary","Cluster","Segment","B2B_flag"]],
        "RFM_Clustering_Result.xlsx",
    )
    add_log("Export Segmentasi")

elif page=="📋 Strategi Pengelolaan":
    st.title("📋 Strategi Pengelolaan Pelanggan")
    if not data_bytes: need_data()
    strat_tab_b2c,strat_tab_b2b=st.tabs(["🏪 Strategi B2C","🏢 Strategi B2B"])

    with strat_tab_b2c:
        st.subheader("Strategi B2C")
        st.caption("Klaster & segmen dihitung ulang sesuai rentang tanggal yang dipilih.")
        df_st=pd.read_excel(io.BytesIO(data_bytes))
        df_st["Waktu Pemesanan"]=pd.to_datetime(df_st["Waktu Pemesanan"],errors="coerce")
        df_st=df_st.dropna(subset=["Waktu Pemesanan"])
        mn_st,mx_st=df_st["Waktu Pemesanan"].min().date(),df_st["Waktu Pemesanan"].max().date()

        sp1,sp2=st.columns(2)
        d1=sp1.date_input("📅 Tanggal Acuan (Baseline):",
            value=max(mn_st,mx_st-timedelta(days=90)),min_value=mn_st,max_value=mx_st,key="stb1")
        sp1.caption("K-Means dari data awal s/d tanggal ini")
        d2=sp2.date_input("📅 Tanggal Saat Ini:",
            value=mx_st,min_value=mn_st,max_value=mx_st,key="stb2")
        sp2.caption("K-Means dari data awal s/d tanggal ini")
        if d1>=d2: st.error("Tanggal Saat Ini harus setelah Baseline."); st.stop()

        with st.spinner(f"⏳ Klaster baseline ({d1})..."):
            p_base=pipeline_b2c_filtered(data_bytes, mn_st, d1)
        with st.spinner(f"⏳ Klaster saat ini ({d2})..."):
            p_cur =pipeline_b2c_filtered(data_bytes, mn_st, d2)
        if not p_base or not p_cur:
            st.warning("Data B2C terlalu sedikit untuk klaster."); st.stop()

        rfm_base=p_base["rfm_b2c"]; rfm_cur=p_cur["rfm_b2c"]
        cl_base=p_base["cl"];        cl_cur=p_cur["cl"]
        segs_base=cl_base["Segment"].tolist()
        segs_cur=cl_cur["Segment"].tolist()
        k_base=p_base["bk"]; k_cur=p_cur["bk"]
        st.info(f"Baseline ({d1}): **{k_base} klaster**, {len(rfm_base):,} B2C  |  Saat Ini ({d2}): **{k_cur} klaster**, {len(rfm_cur):,} B2C", icon="📊")
        st.subheader("📊 Komposisi Klaster — Perbandingan Baseline & Saat Ini")
        def _profil_col(cl, rfm, title):
            st.caption(f"**{title}**")
            cols=["Cluster","Count","Recency","Frequency","Monetary",
                  "R_norm","F_norm","M_norm","R_score","FM_score","Segment"]
            cols=[c for c in cols if c in cl.columns]
            d=cl[cols].copy().rename(columns={"Count":"N","Recency":"Avg R",
                "Frequency":"Avg F","Monetary":"Avg M"})
            d["Avg R"]=d["Avg R"].round(1); d["Avg F"]=d["Avg F"].round(2)
            d["Avg M"]=d["Avg M"].apply(lambda x:f"Rp {x:,.0f}")
            for c in ["R_norm","F_norm","M_norm"]:
                if c in d.columns: d[c]=d[c].round(4)
            aov_d=rfm.groupby("Segment")["Monetary"].mean().round(0).reset_index()
            aov_d.columns=["Segment","AOV"]
            d=d.merge(aov_d,on="Segment",how="left")
            d["AOV"]=d["AOV"].apply(lambda x:f"Rp {x:,.0f}" if pd.notna(x) else "—")
            st.dataframe(d,use_container_width=True,hide_index=True)
            pie=rfm["Segment"].value_counts().reset_index(); pie.columns=["Segment","Count"]
            fig=px.pie(pie,names="Segment",values="Count",color="Segment", color_discrete_map=SEG_COLOR,height=220)
            fig.update_traces(textposition="inside",textinfo="percent+label",textfont_size=9)
            fig.update_layout(showlegend=False,margin=dict(t=10,b=0,l=0,r=0))
            st.plotly_chart(fig,use_container_width=True)

        c_base,c_cur=st.columns(2)
        with c_base: _profil_col(cl_base,rfm_base,f"🔵 Baseline — {d1} ({k_base} klaster)")
        with c_cur:  _profil_col(cl_cur, rfm_cur, f"🟢 Saat Ini — {d2} ({k_cur} klaster)")
        st.subheader(f"Pemetaan Perpindahan Pelanggan Klaster Pelanggan B2C: {d1} → {d2}")
        st.caption(f"Baris = {k_base} klaster Baseline · Kolom = {k_cur} klaster Saat Ini · Nilai = % perpindahan per baris")
        base_uid=rfm_base.set_index("User ID")["Segment"].to_dict()
        cur_uid =rfm_cur.set_index("User ID")["Segment"].to_dict()
        mg=pd.DataFrame([{"User ID":uid,"Seg_Cur":sc,
                           "Seg_Base":base_uid.get(uid,"🆕 New Customer")}
                          for uid,sc in cur_uid.items()])
        r_lab=["🆕 New Customer"]+segs_base
        rs=[r for r in r_lab if r in mg["Seg_Base"].values]
        cs=[c for c in segs_cur if c in mg["Seg_Cur"].values]
        piv=pd.crosstab(mg["Seg_Base"],mg["Seg_Cur"])
        pp=(piv.div(piv.sum(axis=1),axis=0)*100).reindex(index=rs,columns=cs).fillna(0)
        pc=piv.reindex(index=rs,columns=cs).fillna(0)
        fig_h=go.Figure(data=go.Heatmap(z=pp.values, x=cs, y=rs, colorscale=[[0,"#fff5f5"],[0.25,"#fca5a5"],[0.5,"#ef4444"],[0.75,"#b91c1c"],[1.0,"#7f1d1d"]], showscale=True, text=[[f"{pp.loc[r,c]:.0f}%\n({int(pc.loc[r,c])} org)" if (r in pp.index and c in pp.columns and pp.loc[r,c]>0) else "" for c in cs] for r in rs], texttemplate="%{text}", textfont={"size":10}, hovertemplate="Dari: %{y}<br>Ke: %{x}<br>%{z:.1f}% (%{customdata} org)<extra></extra>", customdata=[[int(pc.loc[r,c]) if (r in pc.index and c in pc.columns) else 0 for c in cs] for r in rs]))
        fig_h.update_layout(height=max(280,len(rs)*55), title=f"Transition Matrix Klaster ({d1} → {d2})", xaxis=dict(title=f"Klaster Saat Ini — {d2} ({k_cur} klaster)",tickangle=-25), yaxis=dict(title=f"Klaster Baseline — {d1} ({k_base} klaster)"), margin=dict(t=50,b=90,l=0,r=0))
        st.plotly_chart(fig_h,use_container_width=True)
        st.caption(f"💡 Tiap baris menunjukkan distribusi perpindahan dari klaster baseline ke klaster saat ini. Hover untuk detail. **🆕 New Customer** = pelanggan yang ada di data s/d **{d2}** tapi belum ada di data s/d **{d1}** (bukan hanya yang beli di bulan {d2} saja — mencakup semua pelanggan yang pertama kali muncul setelah {d1}).")

        nc=mg[mg["Seg_Base"]=="🆕 New Customer"]
        repeat_m=mg[(mg["Seg_Base"]!="🆕 New Customer")&(mg["Seg_Cur"]==mg["Seg_Base"])]
        def _rank(s): return SEG_ORDER.index(s) if s in SEG_ORDER else 99
        up=mg[(mg["Seg_Base"]!="🆕 New Customer")&
              (mg["Seg_Cur"]!=mg["Seg_Base"])&
              mg.apply(lambda r:_rank(r["Seg_Cur"])<_rank(r["Seg_Base"]),axis=1)]
        dn=mg[(mg["Seg_Base"]!="🆕 New Customer")&
              (mg["Seg_Cur"]!=mg["Seg_Base"])&
              mg.apply(lambda r:_rank(r["Seg_Cur"])>_rank(r["Seg_Base"]),axis=1)]
        h1,h2,h3,h4=st.columns(4)
        h1.metric("🆕 Baru",f"{len(nc):,}")
        h2.metric("✅ Repeat/Tetap",f"{len(repeat_m):,}")
        h3.metric("⬆️ Upgrade",f"{len(up):,}")
        h4.metric("⬇️ Downgrade",f"{len(dn):,}",delta=f"-{len(dn)}",delta_color="inverse")
        for label,subset,caption in [
            (f"🆕 {len(nc):,} Pelanggan Baru",nc,"Pelanggan yang pertama kali muncul di periode Saat Ini."),
            (f"✅ {len(repeat_m):,} Repeat/Tetap",repeat_m,"Pelanggan lama yang tetap di klaster yang sama."),
            (f"⬆️ {len(up):,} Upgrade Klaster",up,"Pindah ke klaster yang lebih baik — pertahankan!"),
            (f"⬇️ {len(dn):,} Downgrade Klaster",dn,"Pindah ke klaster yang lebih buruk — butuh perhatian!"),
        ]:
            if len(subset)>0:
                with st.expander(f"{label}"):
                    detail_df=subset.merge(rfm_cur[["User ID","Username","Monetary"]].rename(columns={"User ID":"User ID"}),on="User ID",how="left")
                    show_cols=[c for c in ["Username","Seg_Base","Seg_Cur","Monetary"] if c in detail_df.columns]
                    st.dataframe(detail_df[show_cols].reset_index(drop=True), use_container_width=True,height=200)
        st.subheader("👤 Tracking Pelanggan B2C per Klaster")
        trk=mg.merge(rfm_cur[["User ID","Username","Recency","Frequency","Monetary","Segment"]], on="User ID",how="left")
        def _classify_st(r):
            if r["Seg_Base"]=="🆕 New Customer": return "🆕 New"
            if r["Seg_Cur"]==r["Seg_Base"]:      return "✅ Repeat/Tetap"
            return "⬆️ Upgrade" if _rank(r["Seg_Cur"])<_rank(r["Seg_Base"]) else "⬇️ Downgrade"
        trk["Status"]=trk.apply(_classify_st,axis=1)
        tf1,tf2,tf3=st.columns(3)
        stf=tf1.multiselect("Status:",["🆕 New","✅ Repeat/Tetap","⬆️ Upgrade","⬇️ Downgrade"],
                            default=["🆕 New","⬆️ Upgrade","⬇️ Downgrade"])
        sf_=tf2.selectbox("Segmen Saat Ini:",["Semua"]+segs_cur)
        sn=tf3.text_input("Cari Username:",key="stb_s")
        trk_f=trk.copy()
        if stf: trk_f=trk_f[trk_f["Status"].isin(stf)]
        if sf_!="Semua": trk_f=trk_f[trk_f["Seg_Cur"]==sf_]
        if sn: trk_f=trk_f[trk_f["Username"].astype(str).str.contains(sn,case=False,na=False)]
        ts=trk_f[["Username","Seg_Base","Seg_Cur","Status","Recency","Frequency","Monetary"]].reset_index(drop=True)
        ts.index+=1; st.write(f"**{len(ts):,}** pelanggan:"); st.dataframe(ts,use_container_width=True,height=280)
        if not ts.empty: st.download_button("📥 Download Tracking",data=to_xl(ts.reset_index(),"Tracking_Klaster"), file_name=f"Tracking_{d1}_{d2}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.subheader("💡 Strategi Prioritas B2C")
        if "ck" not in st.session_state: st.session_state.ck={}
        focus_segments=[seg for seg in ["At Risk","Loyal Customers"] if seg in segs_cur]
        cc2=st.columns(2)
        for i,seg in enumerate(focus_segments):
            act=STRATEGIES.get(seg,"—"); prio=PRIORITY.get(seg,"low")
            sc_=SEG_COLOR.get(seg,"#888"); cnt=len(rfm_cur[rfm_cur["Segment"]==seg])
            bdg={"high":"🔴 Tinggi","medium":"🟡 Sedang","low":"🟢 Rendah"}[prio]
            with cc2[i%2]:
                st.markdown(f"""<div style="border-left:5px solid {sc_};background:{sc_}12;
border-radius:8px;padding:10px 14px;margin-bottom:6px">
<b style="color:{sc_}">{seg}</b>&nbsp;<small>{bdg}</small><br>
<small style="color:gray">👥 {cnt:,}</small><br><small>💡 {act}</small></div>""",unsafe_allow_html=True)
                ck_k="ck_"+seg.replace(" ","_").replace("'","")
                if ck_k not in st.session_state.ck:
                    st.session_state.ck[ck_k]={}
                ck=st.session_state.ck[ck_k]
                for action_key, action_label in STRATEGY_CHECKLISTS.get(seg, []):
                    current_value=bool(ck.get(action_key, False))
                    ck[action_key]=st.checkbox(action_label, value=current_value, key=f"{ck_k}_{action_key}")
        if not focus_segments:
            st.info("Segmen At Risk atau Loyal Customers belum muncul pada rentang yang dipilih.")
        if st.button("💾 Simpan Progress"):
            ts2=datetime.now().strftime("%Y-%m-%d %H:%M")
            for seg,ck in st.session_state.ck.items():
                seg_name=seg.replace("ck_","").replace("_"," ")
                selected_actions=[action_label for action_key, action_label in STRATEGY_CHECKLISTS.get(seg_name, []) if ck.get(action_key)]
                st.session_state.strat_log.append({"Waktu":ts2, "Segmen":seg_name, "Checklist Selesai":", ".join(selected_actions) if selected_actions else "Belum ada", "Jumlah Aksi":len(selected_actions), "User":st.session_state.uname})
            add_log("Simpan Checklist B2C"); st.success("✅ Tersimpan!")
        if st.session_state.strat_log:
            with st.expander("📋 Riwayat"):
                st.dataframe(pd.DataFrame(st.session_state.strat_log).tail(20), use_container_width=True, height=200)
        st.subheader("💹 Evaluasi Profitabilitas Strategi B2C")
        _ev1, _ev2, _ev3 = st.columns(3)
        _bulan_base_opts = sorted(df_st["Waktu Pemesanan"].dt.to_period("M").unique().astype(str).tolist()); _bulan_eval_opts = _bulan_base_opts.copy()
        _sel_base = _ev1.selectbox("📅 Bulan Klaster (Baseline)", _bulan_base_opts, index=max(0, len(_bulan_base_opts)-2), key="ev_base_month", help="Bulan dimana klaster Loyal/At Risk dihitung (bulan sebelum diskon)")
        _sel_eval = _ev2.selectbox("📅 Bulan Aplikasi Diskon", _bulan_eval_opts, index=len(_bulan_eval_opts)-1, key="ev_eval_month", help="Bulan dimana diskon diberlakukan dan transaksi dihitung")
        _hpp_pct = _ev3.number_input("HPP (%)", min_value=0.0, max_value=100.0, value=41.67, step=0.01, key="ev_hpp_pct", help="Persentase HPP dari revenue. Default 41,67% ≈ Rp5.000/pcs")
        _ec1, _ec2 = st.columns(2)
        _disc_loyal_ev  = _ec1.number_input("Diskon Loyal Customer (%)", 0.0, 100.0, 10.0, 0.5, key="ev_disc_loyal")
        _disc_atrisk_ev = _ec2.number_input("Diskon At Risk (%)",        0.0, 100.0, 20.0, 0.5, key="ev_disc_atrisk")
        if _sel_base >= _sel_eval: st.warning("⚠️ Bulan Aplikasi Diskon harus setelah Bulan Klaster Baseline."); st.stop()
        import calendar
        _base_period = pd.Period(_sel_base, "M"); _eval_period = pd.Period(_sel_eval, "M"); _base_end = _base_period.to_timestamp("M").date(); _eval_start = _eval_period.to_timestamp("S").date(); _eval_end = _eval_period.to_timestamp("M").date()
        with st.spinner(f"⏳ Menghitung klaster baseline ({_sel_base})..."): _p_ev_base = pipeline_b2c_filtered(data_bytes, mn_st, _base_end)
        if not _p_ev_base: st.warning("Data klaster baseline tidak cukup."); st.stop()
        _rfm_base_ev = _p_ev_base["rfm_b2c"]; _base_seg_dict = _rfm_base_ev.set_index("User ID")["Segment"].to_dict()
        _ev_loyal_ids = set(_rfm_base_ev.loc[_rfm_base_ev["Segment"].str.contains("Loyal", case=False, na=False), "User ID"]); _ev_atrisk_ids = set(_rfm_base_ev.loc[_rfm_base_ev["Segment"].str.contains("At Risk", case=False, na=False), "User ID"])
        _df_eval_month = df_st[(df_st["Waktu Pemesanan"].dt.date >= _eval_start) & (df_st["Waktu Pemesanan"].dt.date <= _eval_end)].copy(); _eval_b2c_users = set(df_st.groupby("User ID")["Total Harga Produk"].sum().loc[lambda s: s < B2B_THRESHOLD].index); _df_eval_month = _df_eval_month[_df_eval_month["User ID"].isin(_eval_b2c_users)]
        _n_total_base = len(_rfm_base_ev); _n_not_in_cluster = _n_total_base - len(_ev_loyal_ids) - len(_ev_atrisk_ids)
        st.info(f"🔵 Klaster **{_sel_base}** — Loyal Customer: **{len(_ev_loyal_ids):,}** pelanggan  |  At Risk: **{len(_ev_atrisk_ids):,}** pelanggan  |  Total B2C baseline: **{_n_total_base:,}** pelanggan" + (f"  |  Segmen lain: **{_n_not_in_cluster:,}** pelanggan" if _n_not_in_cluster > 0 else "  |  Semua B2C baseline masuk klaster Loyal/At Risk"), icon="📊")

        _monetary_kum_eval = (
            df_st[df_st["Waktu Pemesanan"].dt.date <= _eval_end]
            .groupby("User ID")["Total Harga Produk"].sum()
        )
        _graduated_to_b2b = {
            uid for uid in set(_rfm_base_ev["User ID"])
            if _monetary_kum_eval.get(uid, 0) >= B2B_THRESHOLD
        }
        if _graduated_to_b2b:
            _grad_detail = _rfm_base_ev[_rfm_base_ev["User ID"].isin(_graduated_to_b2b)][["User ID","Username","Segment"]].copy()
            _grad_detail["Monetary Kumulatif s/d " + str(_sel_eval)] = _grad_detail["User ID"].map(_monetary_kum_eval).apply(lambda v: f"Rp {v:,.0f}")
            st.warning(f"⚠️ **{len(_graduated_to_b2b):,} pelanggan** yang masuk klaster B2C di **{_sel_base}** kini memiliki monetary kumulatif ≥ Rp {B2B_THRESHOLD:,.0f} (threshold B2B) s/d **{_sel_eval}**. Pelanggan ini **tidak lagi masuk program diskon B2C** dan sebaiknya dipindahkan ke segmen B2B pada evaluasi berikutnya.")
            with st.expander(f"👀 Lihat {len(_graduated_to_b2b):,} pelanggan yang pindah B2C → B2B", expanded=False):
                st.dataframe(_grad_detail.reset_index(drop=True), use_container_width=True, hide_index=True)
                st.download_button("📥 Download Daftar Pelanggan Pindah Segmen", data=to_xl(_grad_detail, "Pindah_B2C_ke_B2B", ["User ID"]), file_name=f"Pindah_B2C_B2B_{_sel_base}_{_sel_eval}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        _buyers_eval_month = set(_df_eval_month["User ID"].unique()); _df_base_raw = df_st[(df_st["Waktu Pemesanan"].dt.date <= _base_end) & (df_st["User ID"].isin(_eval_b2c_users))]; _base_all_users = set(_df_base_raw["User ID"].unique()); _repeat_buyers_eval = _buyers_eval_month & _base_all_users
        _repeat_loyal_ids = _repeat_buyers_eval & _ev_loyal_ids; _repeat_atrisk_ids = _repeat_buyers_eval & _ev_atrisk_ids; _repeat_other_ids = _repeat_buyers_eval - _repeat_loyal_ids - _repeat_atrisk_ids
        _df_rep_loyal = _df_eval_month[_df_eval_month["User ID"].isin(_repeat_loyal_ids)]; _df_rep_atrisk = _df_eval_month[_df_eval_month["User ID"].isin(_repeat_atrisk_ids)]; _df_rep_other = _df_eval_month[_df_eval_month["User ID"].isin(_repeat_other_ids)]

        def _seg_agg(df_seg):
            if df_seg.empty: return {"n_cust": 0, "n_tx": 0, "revenue": 0.0}
            return {"n_cust": df_seg["User ID"].nunique(), "n_tx": df_seg["No. Pesanan"].nunique(), "revenue": df_seg["Total Harga Produk"].sum()}

        _agg_loyal  = _seg_agg(_df_rep_loyal)
        _agg_atrisk = _seg_agg(_df_rep_atrisk)
        _agg_other  = _seg_agg(_df_rep_other)
        _cost_loyal_ev       = _agg_loyal["revenue"]  * (_disc_loyal_ev  / 100)
        _cost_atrisk_ev      = _agg_atrisk["revenue"] * (_disc_atrisk_ev / 100)
        _rev_loyal_after_ev  = _agg_loyal["revenue"]  - _cost_loyal_ev
        _rev_atrisk_after_ev = _agg_atrisk["revenue"] - _cost_atrisk_ev

        _total_cust_diskon = _agg_loyal["n_cust"]    + _agg_atrisk["n_cust"]
        _total_tx_diskon   = _agg_loyal["n_tx"]      + _agg_atrisk["n_tx"]
        _total_rev_diskon  = _agg_loyal["revenue"]   + _agg_atrisk["revenue"]
        _total_cost_ev     = _cost_loyal_ev           + _cost_atrisk_ev
        _total_rev_aft_ev  = _total_rev_diskon        - _total_cost_ev
        _rev_bulan_eval   = _df_eval_month["Total Harga Produk"].sum()
        _hpp_bulan_eval   = _rev_bulan_eval * (_hpp_pct / 100)
        _rev_aft_disc_tot = _rev_bulan_eval - _total_cost_ev
        _laba_bersih_ev   = _rev_aft_disc_tot - _hpp_bulan_eval
        _margin_sblm      = 100.0 - _hpp_pct
        _margin_after_ev  = (_laba_bersih_ev / _rev_aft_disc_tot * 100) if _rev_aft_disc_tot else 0.0
        _is_profit_ev     = _laba_bersih_ev > 0
        _n_repeat_eval  = len(_repeat_buyers_eval)
        _n_other_repeat = len(_repeat_other_ids)
        _df_all_kum_eval = df_st[df_st["Waktu Pemesanan"].dt.date <= _eval_end]
        _kum_tx_eval     = _df_all_kum_eval.groupby("User ID")["No. Pesanan"].nunique()
        _rpr_repeat_ids  = {uid for uid in _buyers_eval_month if _kum_tx_eval.get(uid, 0) >= 2}
        _belum_diskon_ids = _rpr_repeat_ids - _base_all_users
        _df_belum_diskon  = _df_eval_month[_df_eval_month["User ID"].isin(_belum_diskon_ids)]
        _agg_belum        = _seg_agg(_df_belum_diskon)
        _new_buyer_ids = _buyers_eval_month - _rpr_repeat_ids
        _df_new_buyers = _df_eval_month[_df_eval_month["User ID"].isin(_new_buyer_ids)]
        _agg_new       = _seg_agg(_df_new_buyers)
        st.subheader(f"📋 Tabel Efektivitas Voucher — {_sel_eval}")
        _voucher_rows = [
            {"Segmen": f"Loyal Customer ({_disc_loyal_ev:.0f}%)", "Jumlah Pelanggan": _agg_loyal["n_cust"], "Jumlah Transaksi": _agg_loyal["n_tx"], "Revenue": f"Rp {_agg_loyal['revenue']:,.0f}", "Total Biaya Diskon": f"Rp {_cost_loyal_ev:,.0f}", "Revenue Setelah Diskon": f"Rp {_rev_loyal_after_ev:,.0f}"},
            {"Segmen": f"At Risk ({_disc_atrisk_ev:.0f}%)", "Jumlah Pelanggan": _agg_atrisk["n_cust"], "Jumlah Transaksi": _agg_atrisk["n_tx"], "Revenue": f"Rp {_agg_atrisk['revenue']:,.0f}", "Total Biaya Diskon": f"Rp {_cost_atrisk_ev:,.0f}", "Revenue Setelah Diskon": f"Rp {_rev_atrisk_after_ev:,.0f}"},
            {"Segmen": f"Repeat Buyer Baru (belum di klaster {_sel_base})", "Jumlah Pelanggan": _agg_belum["n_cust"], "Jumlah Transaksi": _agg_belum["n_tx"], "Revenue": f"Rp {_agg_belum['revenue']:,.0f}", "Total Biaya Diskon": "Rp 0", "Revenue Setelah Diskon": f"Rp {_agg_belum['revenue']:,.0f}"},
            {"Segmen": "Pembeli Baru (1x transaksi)", "Jumlah Pelanggan": _agg_new["n_cust"], "Jumlah Transaksi": _agg_new["n_tx"], "Revenue": f"Rp {_agg_new['revenue']:,.0f}", "Total Biaya Diskon": "Rp 0", "Revenue Setelah Diskon": f"Rp {_agg_new['revenue']:,.0f}"},
        ]
        if _agg_other["n_cust"] > 0:
            _voucher_rows.append({"Segmen": f"Klaster Lain ({_sel_base})", "Jumlah Pelanggan": _agg_other["n_cust"], "Jumlah Transaksi": _agg_other["n_tx"], "Revenue": f"Rp {_agg_other['revenue']:,.0f}", "Total Biaya Diskon": "Rp 0", "Revenue Setelah Diskon": f"Rp {_agg_other['revenue']:,.0f}"})
        _total_cust_tbl = _agg_loyal["n_cust"] + _agg_atrisk["n_cust"] + _agg_belum["n_cust"] + _agg_new["n_cust"] + _agg_other["n_cust"]
        _total_tx_tbl = _agg_loyal["n_tx"] + _agg_atrisk["n_tx"] + _agg_belum["n_tx"] + _agg_new["n_tx"] + _agg_other["n_tx"]
        _voucher_rows.append({"Segmen": "Total", "Jumlah Pelanggan": _total_cust_tbl, "Jumlah Transaksi": _total_tx_tbl, "Revenue": f"Rp {_rev_bulan_eval:,.0f}", "Total Biaya Diskon": f"Rp {_total_cost_ev:,.0f}", "Revenue Setelah Diskon": f"Rp {_rev_aft_disc_tot:,.0f}"})
        _tbl_voucher = pd.DataFrame(_voucher_rows)
        st.dataframe(_tbl_voucher, use_container_width=True, hide_index=True)
        _df_kum_to_eval   = df_st[df_st["Waktu Pemesanan"].dt.date <= _eval_end]
        _monetary_kum_ev  = _df_kum_to_eval.groupby("User ID")["Total Harga Produk"].sum()
        _b2b_users_eval   = set(_monetary_kum_ev[_monetary_kum_ev >= B2B_THRESHOLD].index)
        _buyers_eval_all  = set(df_st[
            (df_st["Waktu Pemesanan"].dt.date >= _eval_start) &
            (df_st["Waktu Pemesanan"].dt.date <= _eval_end)
        ]["User ID"].unique())
        _kum_tx_all       = _df_kum_to_eval.groupby("User ID")["No. Pesanan"].nunique()
        _repeat_b2b_eval  = len({uid for uid in (_buyers_eval_all & _b2b_users_eval) if _kum_tx_all.get(uid, 0) >= 2})
        st.caption(f"💡 Tabel ini hanya mencakup pelanggan **B2C** ({_total_cust_tbl:,} pelanggan). **{_repeat_b2b_eval:,} repeat buyer B2B** tidak masuk program diskon dan tidak ditampilkan di tabel ini (dievaluasi terpisah di tab Strategi B2B). **Repeat Buyer Baru** = {_agg_belum['n_cust']:,} pelanggan B2C dengan ≥2 tx kumulatif namun belum di klaster {_sel_base}. **Pembeli Baru** = {_agg_new['n_cust']:,} pelanggan B2C dengan 1x transaksi.")
        _sm1, _sm2, _sm3 = st.columns(3)
        _sm1.metric(f"💰 Total Pendapatan {_sel_eval}", f"Rp {_rev_bulan_eval:,.0f}")
        _sm2.metric("🏷️ Total Biaya Diskon", f"Rp {_total_cost_ev:,.0f}")
        _sm3.metric("✅ Laba Bersih", f"Rp {_laba_bersih_ev:,.0f}", "✅ Profitable" if _is_profit_ev else "❌ Tidak Profitable", delta_color="normal" if _is_profit_ev else "inverse")
        with st.expander("📋 Lihat Detail Profitabilitas", expanded=False):
            st.subheader(f"📊 Ringkasan Profitabilitas — {_sel_eval}")
            _tbl_profit = pd.DataFrame([
                {"Indikator": f"Total Pendapatan ({_sel_eval})", "Nilai": f"Rp {_rev_bulan_eval:,.0f}", "Keterangan": ""},
                {"Indikator": "Total HPP", "Nilai": f"Rp {_hpp_bulan_eval:,.0f}", "Keterangan": f"Rp5.000/pcs atau {_hpp_pct:.2f}%"},
                {"Indikator": "Total Biaya Diskon", "Nilai": f"Rp {_total_cost_ev:,.0f}", "Keterangan": "Gabungan Loyal + At Risk"},
                {"Indikator": "Laba Bersih", "Nilai": f"Rp {_laba_bersih_ev:,.0f}", "Keterangan": ""},
                {"Indikator": "Margin Laba Bersih Sebelum Diskon", "Nilai": f"{_margin_sblm:.2f}%", "Keterangan": ""},
                {"Indikator": "Margin Laba Bersih Setelah Diskon", "Nilai": f"{_margin_after_ev:.2f}%", "Keterangan": ""},
                {"Indikator": "Status Profitabilitas", "Nilai": "✅ Profitable" if _is_profit_ev else "❌ Tidak Profitable", "Keterangan": ""},
            ])
            st.dataframe(_tbl_profit, use_container_width=True, hide_index=True)
            _bc_segs = [f"Loyal Customer ({_disc_loyal_ev:.0f}%)", f"At Risk ({_disc_atrisk_ev:.0f}%)"]
            fig_ev = go.Figure()
            fig_ev.add_trace(go.Bar(name="Revenue Sebelum Diskon", x=_bc_segs, y=[_agg_loyal["revenue"], _agg_atrisk["revenue"]], marker_color="#3b82f6", text=[f"Rp {v:,.0f}" for v in [_agg_loyal["revenue"], _agg_atrisk["revenue"]]], textposition="outside"))
            fig_ev.add_trace(go.Bar(name="Revenue Setelah Diskon", x=_bc_segs, y=[_rev_loyal_after_ev, _rev_atrisk_after_ev], marker_color="#22c55e", text=[f"Rp {v:,.0f}" for v in [_rev_loyal_after_ev, _rev_atrisk_after_ev]], textposition="outside"))
            fig_ev.add_trace(go.Bar(name="Total Biaya Diskon", x=_bc_segs, y=[_cost_loyal_ev, _cost_atrisk_ev], marker_color="#ef4444", text=[f"Rp {v:,.0f}" for v in [_cost_loyal_ev, _cost_atrisk_ev]], textposition="outside"))
            fig_ev.update_layout(barmode="group", height=400, legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left"), xaxis=dict(gridcolor="#eee"), yaxis=dict(title="Nilai (Rp)", tickformat=",.0f", gridcolor="#eee"), margin=dict(t=60, b=40, l=0, r=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_ev, use_container_width=True)
            st.subheader(f"👥 Tabel Detail Pelanggan yang Berhasil Dipertahankan — {_sel_eval}")

            def _build_retained(df_seg, label):
                if df_seg.empty: return pd.DataFrame()
                agg = df_seg.groupby("User ID").agg(Nama_Pelanggan=("Username (Pembeli)", "first"), Jumlah_Transaksi=("No. Pesanan", "nunique"), Revenue=("Total Harga Produk", "sum")).reset_index()
                agg["Segmen Bulan Sebelumnya"] = label
                return agg

            _retained_loyal = _build_retained(_df_rep_loyal, f"Loyal Customer ({_sel_base})"); _retained_atrisk = _build_retained(_df_rep_atrisk, f"At Risk ({_sel_base})"); _retained_other = _build_retained(_df_rep_other, f"Segmen Lain ({_sel_base})"); _retained_all = pd.concat([_retained_loyal, _retained_atrisk, _retained_other], ignore_index=True)
            if not _retained_all.empty:
                _ret_disp = _retained_all[["User ID", "Nama_Pelanggan", "Segmen Bulan Sebelumnya", "Jumlah_Transaksi", "Revenue"]].copy()
                _ret_disp = _ret_disp.rename(columns={"Nama_Pelanggan": "Nama Pelanggan", "Jumlah_Transaksi": f"Jumlah Transaksi {_sel_eval}", "Revenue": f"Revenue {_sel_eval}"})
                _ret_disp[f"Revenue {_sel_eval}"] = _ret_disp[f"Revenue {_sel_eval}"].apply(lambda v: f"Rp {v:,.0f}")
                _ret_disp = _ret_disp.sort_values("Segmen Bulan Sebelumnya").reset_index(drop=True)
                _ret_disp.index += 1
                st.dataframe(_ret_disp, use_container_width=True, height=360)
                st.download_button("📥 Download Detail Pelanggan", data=to_xl(_retained_all, "Detail_Retained", ["User ID"]), file_name=f"Retained_{_sel_base}_{_sel_eval}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info(f"Tidak ada pembeli berulang pada {_sel_eval}.")

            st.caption(f"💡 Metode: (1) Cari pembeli berulang {_sel_eval} = irisan pembeli bulan ini dengan pelanggan baseline. (2) Cocokkan klaster {_sel_base}: Loyal → diskon {_disc_loyal_ev:.0f}%, At Risk → diskon {_disc_atrisk_ev:.0f}%. (3) HPP = {_hpp_pct:.2f}% × Total Revenue seluruh {_sel_eval}.")
        st.markdown("---")

        st.subheader("📁 Perancangan Strategi")
        up_strat=st.file_uploader("Upload dokumen strategi",
            type=["pdf","docx","xlsx","pptx","csv","txt"],accept_multiple_files=True,key="strat_doc")
        if up_strat and _MODULES_OK:
            for f in up_strat:
                try: save_strategy_doc(f.read(),f.name); st.success(f"✅ {f.name}")
                except Exception as e: st.warning(f"Gagal: {e}")
        if _MODULES_OK:
            docs=list_strategy_docs()
            if docs: st.dataframe(pd.DataFrame(docs),use_container_width=True,hide_index=True)

    with strat_tab_b2b:
        st.subheader("Strategi Pengelolaan Pelanggan B2B")
        p3=pipeline(data_bytes); rfm_b2b3=p3["rfm_b2b"]
        df_raw3=pd.read_excel(io.BytesIO(data_bytes)); df_raw3["Waktu Pemesanan"]=pd.to_datetime(df_raw3["Waktu Pemesanan"],errors="coerce"); df_raw3=df_raw3.dropna(subset=["Waktu Pemesanan"]); _b2b_mn=df_raw3["Waktu Pemesanan"].min().date(); _b2b_mx=df_raw3["Waktu Pemesanan"].max().date()
        _bv1,_bv2=st.columns(2)
        _b2b_d1=_bv1.date_input("📅 Dari",value=_b2b_mn,min_value=_b2b_mn,max_value=_b2b_mx,key="b2b_ev_d1")
        _b2b_d2=_bv2.date_input("📅 Sampai",value=_b2b_mx,min_value=_b2b_mn,max_value=_b2b_mx,key="b2b_ev_d2")
        if _b2b_d1>_b2b_d2: st.error("Tanggal Dari harus sebelum Sampai."); st.stop()

        df_b2b_f=df_raw3[(df_raw3["Waktu Pemesanan"].dt.date>=_b2b_d1)&
                         (df_raw3["Waktu Pemesanan"].dt.date<=_b2b_d2)].copy()
        _df_kum_b2b = df_raw3[df_raw3["Waktu Pemesanan"].dt.date <= _b2b_d2]
        _monetary_kum = _df_kum_b2b.groupby("User ID")["Total Harga Produk"].sum()
        _b2b_user_ids = set(_monetary_kum[_monetary_kum >= B2B_THRESHOLD].index)
        _df_kum_before = df_raw3[df_raw3["Waktu Pemesanan"].dt.date < _b2b_d1]
        _monetary_kum_before = _df_kum_before.groupby("User ID")["Total Harga Produk"].sum()
        _snap_b2b=pd.Timestamp(_b2b_d2)+pd.Timedelta(days=1)
        rfm_b2b_ev=build_rfm(df_b2b_f[df_b2b_f["User ID"].isin(_b2b_user_ids)], _snap_b2b, {"First_Purchase": ("Waktu Pemesanan", "min"), "Last_Purchase": ("Waktu Pemesanan", "max"), "Kanal": ("Kanal", "last")})
        rfm_b2b_ev["Monetary_Kum_Before"] = rfm_b2b_ev["User ID"].map(lambda uid: _monetary_kum_before.get(uid, 0))
        rfm_b2b_ev["Monetary_Kum"] = rfm_b2b_ev["User ID"].map(_monetary_kum)
        _first_tx_all = df_raw3.groupby("User ID")["Waktu Pemesanan"].min().dt.date; rfm_b2b_ev["First_Ever"] = rfm_b2b_ev["User ID"].map(_first_tx_all)
        rfm_b2b_ev["Tipe"] = rfm_b2b_ev["User ID"].apply(lambda uid: "🆕 B2B Baru" if _monetary_kum_before.get(uid, 0) < B2B_THRESHOLD else "🔁 B2B Lama")
        _b2b_lama=rfm_b2b_ev[rfm_b2b_ev["Tipe"]=="🔁 B2B Lama"]; _b2b_baru=rfm_b2b_ev[rfm_b2b_ev["Tipe"]=="🆕 B2B Baru"]
        _kum_tx_b2b     = _df_kum_b2b.groupby("User ID")["No. Pesanan"].nunique()
        _b2b_buyers_period = set(df_b2b_f["User ID"].unique()) & _b2b_user_ids
        _n_b2b_repeat   = int(sum(1 for uid in _b2b_buyers_period if _kum_tx_b2b.get(uid, 0) >= 2))
        _n_b2b_new_1x   = len(_b2b_buyers_period) - _n_b2b_repeat
        st.markdown("---")
        _k1,_k2,_k3,_k4=st.columns(4)
        _k1.metric("🏢 Total B2B Aktif",f"{len(rfm_b2b_ev):,}", help="Pelanggan B2B yang bertransaksi dalam rentang ini")
        _k2.metric("🔁 B2B Lama",f"{len(_b2b_lama):,}", help="First purchase sebelum rentang ini")
        _k3.metric("🆕 B2B Baru",f"{len(_b2b_baru):,}", help="First purchase dalam rentang ini")
        _k4.metric("💰 Total Revenue B2B", f"Rp {rfm_b2b_ev['Monetary'].sum()/1e6:.1f}jt")
        _k5,_k6=st.columns(2)
        _k5.metric("🔄 B2B Berulang (≥2 tx kumulatif)",f"{_n_b2b_repeat:,}", help="Pelanggan B2B dengan ≥2 transaksi kumulatif s/d akhir rentang — sama dengan definisi RPR di Home")
        _k6.metric("1️⃣ B2B Beli 1x",f"{_n_b2b_new_1x:,}", help="Pelanggan B2B yang kumulatifnya baru 1 transaksi — belum dihitung sebagai berulang")
        _n_baru_repeat = int((_b2b_baru["User ID"].map(_kum_tx_b2b) >= 2).sum())
        st.caption(f"💡 **B2B Berulang** ({_n_b2b_repeat:,}) harus sama dengan kolom **↳ B2B Berulang** di tabel RPR Home pada bulan yang sama. B2B Baru ({len(_b2b_baru):,}) + B2B Lama ({len(_b2b_lama):,}) = Total B2B Aktif ({len(rfm_b2b_ev):,}). **B2B Baru** = pelanggan yang belum mencapai threshold B2B sebelum rentang ini (termasuk pelanggan lama B2C yang baru naik ke B2B). Dari {len(_b2b_baru):,} B2B Baru, yang sudah berulang (≥2 tx kumulatif): **{_n_baru_repeat:,}** pelanggan.")
        st.markdown("---")
        st.subheader("📋 Tabel Pelanggan B2B — Lama & Baru")
        rfm_b2b_ev["Berulang"] = rfm_b2b_ev["User ID"].apply(lambda uid: "🔄 Berulang" if _kum_tx_b2b.get(uid, 0) >= 2 else "1️⃣ Beli 1x")
        _tbl_b2b_summary=pd.DataFrame([
            {"Tipe": "🔁 B2B Lama", "Jumlah Pelanggan": len(_b2b_lama), "Berulang (≥2 tx)": int((_b2b_lama["User ID"].map(_kum_tx_b2b) >= 2).sum()), "Beli 1x": int((_b2b_lama["User ID"].map(_kum_tx_b2b) < 2).sum()), "Total Transaksi": int(_b2b_lama["Frequency"].sum()), "Total Revenue": f"Rp {_b2b_lama['Monetary'].sum():,.0f}", "Avg Monetary": f"Rp {_b2b_lama['Monetary'].mean():,.0f}" if len(_b2b_lama) else "Rp 0", "Avg Recency (hari)": round(_b2b_lama["Recency"].mean(),1) if len(_b2b_lama) else 0},
            {"Tipe": "🆕 B2B Baru", "Jumlah Pelanggan": len(_b2b_baru), "Berulang (≥2 tx)": int((_b2b_baru["User ID"].map(_kum_tx_b2b) >= 2).sum()), "Beli 1x": int((_b2b_baru["User ID"].map(_kum_tx_b2b) < 2).sum()), "Total Transaksi": int(_b2b_baru["Frequency"].sum()), "Total Revenue": f"Rp {_b2b_baru['Monetary'].sum():,.0f}", "Avg Monetary": f"Rp {_b2b_baru['Monetary'].mean():,.0f}" if len(_b2b_baru) else "Rp 0", "Avg Recency (hari)": round(_b2b_baru["Recency"].mean(),1) if len(_b2b_baru) else 0},
            {"Tipe": "Total", "Jumlah Pelanggan": len(rfm_b2b_ev), "Berulang (≥2 tx)": _n_b2b_repeat, "Beli 1x": _n_b2b_new_1x, "Total Transaksi": int(rfm_b2b_ev["Frequency"].sum()), "Total Revenue": f"Rp {rfm_b2b_ev['Monetary'].sum():,.0f}", "Avg Monetary": f"Rp {rfm_b2b_ev['Monetary'].mean():,.0f}" if len(rfm_b2b_ev) else "Rp 0", "Avg Recency (hari)": round(rfm_b2b_ev["Recency"].mean(),1) if len(rfm_b2b_ev) else 0},
        ])
        st.dataframe(_tbl_b2b_summary,use_container_width=True,hide_index=True)
        _all_b2b_ever = set(_monetary_kum[_monetary_kum >= B2B_THRESHOLD].index)
        _n_b2b_ever, _n_b2b_active = len(_all_b2b_ever), len(rfm_b2b_ev)  # aktif dalam rentang
        _n_b2b_inactive = _n_b2b_ever - _n_b2b_active  # pernah B2B tapi tidak transaksi di rentang
        st.markdown("##### 👥 Keseluruhan Pelanggan B2B (Sepanjang Data)")
        _tot1, _tot2, _tot3 = st.columns(3)
        _tot1.metric("📊 Total Pelanggan B2B (Kumulatif)", f"{_n_b2b_ever:,}", help="Seluruh pelanggan yang pernah mencapai threshold B2B s/d akhir rentang")
        _tot2.metric("✅ Aktif dalam Rentang", f"{_n_b2b_active:,}", help="Pelanggan B2B yang bertransaksi dalam rentang yang dipilih")
        _tot3.metric("💤 Tidak Aktif dalam Rentang", f"{_n_b2b_inactive:,}", help="Pelanggan B2B yang pernah beli tapi tidak transaksi dalam rentang ini")
        _rfm_b2b_all = build_rfm(df_raw3[df_raw3["User ID"].isin(_all_b2b_ever)], _snap_b2b, {"First_Purchase": ("Waktu Pemesanan", "min"), "Last_Purchase": ("Waktu Pemesanan", "max"), "Kanal": ("Kanal", "last")})
        _rfm_b2b_all["Aktif_Rentang"] = _rfm_b2b_all["User ID"].apply(lambda uid: "✅ Aktif" if uid in set(rfm_b2b_ev["User ID"]) else "💤 Tidak Aktif")
        _rfm_b2b_all["Tipe"] = _rfm_b2b_all["User ID"].apply(lambda uid: "🆕 B2B Baru" if _monetary_kum_before.get(uid, 0) < B2B_THRESHOLD else "🔁 B2B Lama")
        with st.expander(f"👥 Detail Seluruh Pelanggan B2B Aktif — {len(rfm_b2b_ev):,} pelanggan",expanded=False):
            _b2b_detail=rfm_b2b_ev[["User ID","Username","Tipe","Recency","Frequency","Monetary","Monetary_Kum","First_Ever","Last_Purchase","Kanal"]].copy()
            _b2b_detail=_b2b_detail.sort_values(["Tipe","Monetary"],ascending=[True,False]).reset_index(drop=True)
            _b2b_detail.index+=1
            _b2b_detail["Monetary"]=_b2b_detail["Monetary"].apply(lambda x:f"Rp {x:,.0f}")
            _b2b_detail["Monetary_Kum"]=_b2b_detail["Monetary_Kum"].apply(lambda x:f"Rp {x:,.0f}")
            _b2b_detail=_b2b_detail.rename(columns={"Monetary_Kum":"Monetary Kumulatif (s/d akhir rentang)"})
            st.dataframe(_b2b_detail,use_container_width=True,height=320)
        with st.expander(f"📋 Tabel RFM Lengkap Semua Pelanggan B2B — {_n_b2b_ever:,} pelanggan (termasuk tidak aktif)",expanded=False):
            _rfm_disp = _rfm_b2b_all[["User ID","Username","Tipe","Aktif_Rentang","Recency","Frequency","Monetary","First_Purchase","Last_Purchase","Kanal"]].copy().sort_values(["Aktif_Rentang","Monetary"],ascending=[True,False]).reset_index(drop=True)
            _rfm_disp.index += 1
            _avg_r, _avg_f, _avg_m = _rfm_disp["Recency"].mean(), _rfm_disp["Frequency"].mean(), _rfm_disp["Monetary"].mean()
            _ac1,_ac2,_ac3 = st.columns(3)
            _ac1.metric("Avg Recency (hari)", f"{_avg_r:.1f}")
            _ac2.metric("Avg Frequency (tx)", f"{_avg_f:.2f}")
            _ac3.metric("Avg Monetary", f"Rp {_avg_m:,.0f}")
            _rfm_disp_show = _rfm_disp.copy(); _rfm_disp_show["Monetary"] = _rfm_disp_show["Monetary"].apply(lambda x: f"Rp {x:,.0f}")
            st.dataframe(_rfm_disp_show, use_container_width=True, height=360)
            _export_b2b = _rfm_b2b_all[["User ID","Username","Tipe","Aktif_Rentang","Recency","Frequency","Monetary","First_Purchase","Last_Purchase","Kanal"]].copy().sort_values(["Aktif_Rentang","Monetary"],ascending=[True,False])
            st.download_button("📥 Download Tabel RFM Lengkap B2B", data=to_xl(_export_b2b, "RFM_B2B_Lengkap", ["User ID"]), file_name=f"RFM_B2B_Lengkap_{_b2b_d1}_{_b2b_d2}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        maybe_sync_download(
            st.download_button("📥 Download Data B2B Aktif", data=to_xl(rfm_b2b_ev.sort_values("Monetary",ascending=False),"Data_B2B",["User ID"]), file_name="Data_B2B.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            rfm_b2b_ev.sort_values("Monetary",ascending=False),
            "Data_B2B.xlsx",
        )
        st.markdown("---")
        df_fp=df_raw3.groupby("User ID").agg(Total_Monetary=("Total Harga Produk","sum"), First_Purchase=("Waktu Pemesanan","min")).reset_index()
        df_b2b_acq=df_fp[df_fp["Total_Monetary"]>=B2B_THRESHOLD].copy()
        df_b2b_acq["Bulan"]=df_b2b_acq["First_Purchase"].dt.to_period("M").astype(str)
        acq_m=df_b2b_acq.groupby("Bulan").size().reset_index(name="B2B_Baru")
        fig_acq=px.bar(acq_m,x="Bulan",y="B2B_Baru", title="Akuisisi Pelanggan B2B Baru per Bulan", labels={"B2B_Baru":"Pelanggan Baru"})
        fig_acq.update_layout(height=280,xaxis_tickangle=-30,margin=dict(t=40,b=60,l=0,r=0))
        fig_acq.update_traces(hovertemplate="<b>%{x}</b><br>B2B Baru: %{y}<extra></extra>")
        st.plotly_chart(fig_acq,use_container_width=True)

elif page=="📥 Data & Input":
    st.title("📥 Data & Input Transaksi")
    if _MODULES_OK:
        st.caption(f"📁 Data disimpan di: `{PROCESSED_DIR}`")

    sumber=st.radio("Pilih sumber data:", ["🟠 Shopee","⚫ TikTok Shop","💬 WhatsApp Business"], horizontal=True)
    if "WhatsApp" in sumber:
        st.info("Format file WA Business (5 kolom wajib): `No. Pesanan` · `User ID` · `Waktu Pesanan Dibuat` · `Username (Pembeli)` · `Total Harga Produk`",icon="ℹ️")
        up_wa=st.file_uploader("Upload file WA Business (.xlsx/.csv)", type=["xlsx","csv"],accept_multiple_files=True,key="up_wa")
        if up_wa:
            frames_wa=[]
            for f in up_wa:
                try:
                    fb=f.read(); fn=f.name.lower()
                    if fn.endswith(".csv"):
                        dwa=pd.read_csv(io.BytesIO(fb),dtype={"No. Pesanan":str,"User ID":str})
                    else:
                        dwa=pd.read_excel(io.BytesIO(fb),dtype={"No. Pesanan":str,"User ID":str})
                    dwa.columns=dwa.columns.str.strip()
                    dwa=dwa.loc[:,~dwa.columns.duplicated()].copy()
                    for wc in ["Waktu Pesanan Dibuat","Waktu Pemesanan","Created Time"]:
                        if wc in dwa.columns and "Waktu Pemesanan" not in dwa.columns:
                            dwa=dwa.rename(columns={wc:"Waktu Pemesanan"}); break
                    dwa["Kanal"]="WhatsApp"
                    dwa["No. Pesanan"]=dwa["No. Pesanan"].astype(str).str.strip()
                    frames_wa.append(dwa)
                    st.success(f"✅ {f.name} — {len(dwa):,} baris")
                except Exception as e:
                    st.warning(f"⚠️ {f.name}: {e}")
            if frames_wa:
                df_wa=pd.concat(frames_wa,ignore_index=True)
                st.write(f"**{len(df_wa):,}** baris siap"); st.dataframe(df_wa.head(10),use_container_width=True)
                df_wa["Waktu Pemesanan"]=pd.to_datetime(df_wa["Waktu Pemesanan"],errors="coerce")
                for c in REQUIRED_COLS:
                    if c not in df_wa.columns: df_wa[c]=None
                df_exist_wa=None
                if data_bytes: df_exist_wa=pd.read_excel(io.BytesIO(data_bytes),dtype={"No. Pesanan":str,"User ID":str})
                if _MODULES_OK and df_exist_wa is not None:
                    df_mg_wa,n_add_wa,n_dup_wa=merge_to_gabungan(df_wa,df_exist_wa)
                elif df_exist_wa is not None:
                    df_exist_wa["No. Pesanan"]=df_exist_wa["No. Pesanan"].astype(str).str.strip()
                    dup=df_wa["No. Pesanan"].isin(df_exist_wa["No. Pesanan"])
                    n_dup_wa=int(dup.sum()); n_add_wa=int((~dup).sum())
                    df_mg_wa=pd.concat([df_exist_wa,df_wa[~dup]],ignore_index=True)
                else:
                    df_mg_wa=df_wa; n_add_wa=len(df_wa); n_dup_wa=0
                wa1,wa2,wa3=st.columns(3)
                wa1.metric("🔁 Duplikat",f"{n_dup_wa:,}")
                wa2.metric("✅ Baru",f"{n_add_wa:,}")
                wa3.metric("📊 Total",f"{len(df_mg_wa):,}")
                st.download_button("📥 Download Transaksi_Gabungan.xlsx", data=to_xl(df_mg_wa,"Transaksi_Gabungan",["No. Pesanan","User ID"]), file_name="Transaksi_Gabungan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
                if _MODULES_OK and st.button("💾 Simpan ke Storage (WA)",type="primary"):
                    try:
                        save_gabungan(df_mg_wa); st.cache_data.clear()
                        add_log("Upload WA Business",f"{n_add_wa:,} baris baru"); st.success("✅ Tersimpan!"); st.rerun()
                    except Exception as e: st.error(f"Gagal: {e}")
    else:
        kanal_str="Shopee" if "Shopee" in sumber else "TikTok"
        cur_y,cur_m=datetime.now().year,datetime.now().month
        uc1,uc2=st.columns(2)
        tahun_up=uc1.number_input("Tahun data:",min_value=2020,max_value=2030,value=cur_y)
        bulan_up=uc2.selectbox("Bulan data:",list(range(1,13)), format_func=lambda x:f"{x} — {['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'][x-1]}", index=cur_m-1)
        st.caption(f"File akan disimpan sebagai: **{tahun_up}_{bulan_up}_{kanal_str}.xlsx**")
        if kanal_str=="Shopee":
            st.info("Format Shopee: `No. Pesanan` · `Waktu Pesanan Dibuat` · `Username (Pembeli)` · `Jumlah Produk di Pesan` · `Harga Awal Produk` · `Status Pesanan`",icon="ℹ️")
        else:
            st.info("Format TikTok: `Order ID` · `Created Time` · `Buyer Username` · `Quantity` · `SKU Unit Original Price` · `Order Status`",icon="ℹ️")
        up_files=st.file_uploader(f"Upload {kanal_str} (.xlsx/.xls/.csv)", type=["xlsx","xls","csv"],accept_multiple_files=True,key=f"up_{kanal_str}")
        df_processed=None
        if up_files:
            frames=[]; errors=[]
            prog=st.progress(0)
            for idx,f in enumerate(up_files):
                fb=f.read()
                try:
                    if _MODULES_OK:
                        proc_fn=preprocess_shopee if kanal_str=="Shopee" else preprocess_tiktok
                        df_f=proc_fn(fb,f.name)
                    else:
                        try:
                            fn_lower=f.name.lower()
                            if fn_lower.endswith(".csv"):
                                df_f=pd.read_csv(io.BytesIO(fb),dtype=str,encoding="utf-8-sig",on_bad_lines="skip")
                            else:
                                df_f=pd.read_excel(io.BytesIO(fb),dtype=str)
                        except Exception:
                            df_f=pd.read_excel(io.BytesIO(fb))
                        df_f.columns=df_f.columns.str.strip()
                        df_f=df_f.loc[:,~df_f.columns.duplicated()].copy()
                        if kanal_str=="TikTok":
                            for old,new in [("Order ID","No. Pesanan"),("Created Time","Waktu Pemesanan"),
                                            ("Buyer Username","Username (Pembeli)"),("Quantity","Jumlah"),
                                            ("SKU Unit Original Price","Harga Awal")]:
                                if old in df_f.columns: df_f=df_f.rename(columns={old:new})
                        else:
                            for old,new in [("Waktu Pesanan Dibuat","Waktu Pemesanan"),
                                            ("Jumlah Produk di Pesan","Jumlah"),
                                            ("Harga Awal Produk","Harga Awal")]:
                                if old in df_f.columns: df_f=df_f.rename(columns={old:new})
                        if "Total Harga Produk" not in df_f.columns:
                            if "Jumlah" in df_f.columns and "Harga Awal" in df_f.columns:
                                df_f["Total Harga Produk"]=(pd.to_numeric(df_f["Jumlah"],errors="coerce").fillna(0)*pd.to_numeric(df_f["Harga Awal"],errors="coerce").fillna(0))
                            elif "Order Amount" in df_f.columns:
                                df_f["Total Harga Produk"]=pd.to_numeric(df_f["Order Amount"],errors="coerce").fillna(0)
                        if "User ID" not in df_f.columns and "Username (Pembeli)" in df_f.columns:
                            prefix="Shope" if kanal_str=="Shopee" else "TikTok"
                            try:
                                codes=pd.factorize(df_f["Username (Pembeli)"].fillna("").astype(str))[0]
                                df_f["User ID"]=prefix+"_"+(pd.Series(codes)+1).astype(str).str.zfill(5)
                            except Exception:
                                df_f["User ID"]=prefix+"_"+(pd.RangeIndex(len(df_f))+1).astype(str).str.zfill(5)
                        if "No. Pesanan" in df_f.columns:
                            df_f["No. Pesanan"]=df_f["No. Pesanan"].astype(str).str.strip()
                        df_f["Kanal"]=kanal_str
                    if df_f is not None and not df_f.empty:
                        frames.append(df_f); st.toast(f"✅ {f.name} — {len(df_f):,} baris")
                        if _MODULES_OK:
                            try: save_raw_file(fb,kanal_str,int(tahun_up),int(bulan_up))
                            except Exception: pass
                        add_log(f"Upload {kanal_str}",f.name)
                except Exception as e:
                    errors.append(f"{f.name}: {str(e)[:100]}")
                prog.progress((idx+1)/len(up_files))
            for err in errors: st.warning(f"⚠️ {err}")
            if frames:
                df_processed=pd.concat(frames,ignore_index=True)
                st.success(f"✅ {len(df_processed):,} baris dari {len(frames)} file")
                with st.expander("Preview data"): st.dataframe(df_processed.head(10),use_container_width=True)

        if df_processed is not None and not df_processed.empty:
            st.markdown("---")
            st.subheader("🔗 Merge")
            for wc in ["Waktu Pesanan Dibuat","Created Time"]:
                if wc in df_processed.columns and "Waktu Pemesanan" not in df_processed.columns:
                    df_processed=df_processed.rename(columns={wc:"Waktu Pemesanan"})
            df_processed["Waktu Pemesanan"]=pd.to_datetime(df_processed["Waktu Pemesanan"],errors="coerce")
            if "No. Pesanan" not in df_processed.columns: df_processed["No. Pesanan"]="UNK_"+pd.RangeIndex(len(df_processed)).astype(str)
            df_processed["No. Pesanan"]=df_processed["No. Pesanan"].astype(str).str.strip()
            for c in REQUIRED_COLS:
                if c not in df_processed.columns: df_processed[c]=None
            df_exist=None
            if data_bytes: df_exist=pd.read_excel(io.BytesIO(data_bytes),dtype={"No. Pesanan":str,"User ID":str})
            if _MODULES_OK:
                df_mg,n_add,n_dup=merge_to_gabungan(df_processed,df_exist)
            else:
                if df_exist is not None:
                    df_exist["No. Pesanan"]=df_exist["No. Pesanan"].astype(str).str.strip()
                    dup_mask=df_processed["No. Pesanan"].isin(df_exist["No. Pesanan"])
                    n_dup=int(dup_mask.sum()); n_add=int((~dup_mask).sum())
                    df_mg=pd.concat([df_exist,df_processed[~dup_mask]],ignore_index=True)
                else:
                    n_dup=0; n_add=len(df_processed); df_mg=df_processed.copy()
            m1,m2,m3=st.columns(3)
            m1.metric("🔁 Duplikat skip",f"{n_dup:,}")
            m2.metric("✅ Baris Baru",f"{n_add:,}")
            m3.metric("📊 Total",f"{len(df_mg):,}")
            sa1,sa2=st.columns(2)
            with sa1:
                xl=to_xl(df_mg,"Transaksi_Gabungan",["No. Pesanan","User ID"])
                if _MODULES_OK:
                    if st.button("💾 Simpan ke Storage",type="primary"):
                        try:
                            save_gabungan(df_mg); st.cache_data.clear()
                            add_log("Simpan Gabungan",f"{len(df_mg):,} baris"); st.success("✅ Tersimpan!"); st.rerun()
                        except Exception as e: st.error(f"Gagal: {e}")
                st.download_button("📥 Download Transaksi_Gabungan.xlsx", data=xl, file_name="Transaksi_Gabungan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with sa2:
                st.markdown("**Update manual:**\n1. Download → 2. Timpa file lokal → 3. Tekan **R**")
    st.markdown("---")
    st.subheader("📊 Monitoring Data Transaksi")
    if not data_bytes: st.info("Belum ada data aktif.",icon="📂"); st.stop()
    dr=pd.read_excel(io.BytesIO(data_bytes),dtype={"No. Pesanan":str,"User ID":str})
    dr["Waktu Pemesanan"]=pd.to_datetime(dr["Waktu Pemesanan"],errors="coerce")
    ic1,ic2,ic3,ic4=st.columns(4)
    ic1.metric("📦 Total Transaksi",f"{len(dr):,}")
    ic2.metric("📅 Terlama",str(dr["Waktu Pemesanan"].min().date()))
    ic3.metric("📅 Terbaru",str(dr["Waktu Pemesanan"].max().date()))
    ic4.metric("👥 Jumlah Pelanggan Keseluruhan",f"{dr['User ID'].nunique():,}")
    ff1,ff2,ff3,ff4=st.columns(4)
    kf2=ff1.selectbox("Kanal:",["Semua"]+sorted(dr["Kanal"].unique()),key="mk")
    su2=ff2.text_input("Cari Username:",key="mu")
    dr_mn,dr_mx=dr["Waktu Pemesanan"].min().date(),dr["Waktu Pemesanan"].max().date()
    d_s=ff3.date_input("Dari:",value=dr_mn,min_value=dr_mn,max_value=dr_mx,key="ms")
    d_e=ff4.date_input("Sampai:",value=dr_mx,min_value=dr_mn,max_value=dr_mx,key="me")
    dsh=dr.copy()
    if kf2!="Semua": dsh=dsh[dsh["Kanal"]==kf2]
    if su2: dsh=dsh[dsh["Username (Pembeli)"].astype(str).str.contains(su2,case=False,na=False)]
    dsh=dsh[(dsh["Waktu Pemesanan"].dt.date>=d_s)&(dsh["Waktu Pemesanan"].dt.date<=d_e)]
    st.write(f"**{len(dsh):,}** transaksi:")
    st.dataframe(dsh,use_container_width=True,height=320)
    dc1,dc2=st.columns(2)
    with dc1:
        kp=dsh.groupby("Kanal").size().reset_index(name="Transaksi")
        fig_kp=px.pie(kp,names="Kanal",values="Transaksi",color="Kanal",color_discrete_map=KANAL_COLOR,title="Transaksi per Kanal")
        fig_kp.update_layout(height=240,margin=dict(t=40,b=0,l=0,r=0))
        st.plotly_chart(fig_kp,use_container_width=True)
    with dc2:
        maybe_sync_download(
            st.download_button("📥 Export", data=to_xl(dsh,"Data_Transaksi"),file_name="Monitoring_Transaksi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            dsh,
            "Monitoring_Transaksi.xlsx",
        )
        add_log("Export Monitoring",f"{len(dsh):,} baris")