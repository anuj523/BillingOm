"""
IDBI AWS Billing Tool  v3.0
Runs locally — nothing leaves this machine.

Usage:  streamlit run app.py
"""
import os, io, tempfile, traceback
import streamlit as st

st.set_page_config(
    page_title="IDBI AWS Billing Tool",
    page_icon="📊",
    layout="wide",
)

# ── Shared helpers ─────────────────────────────────────────────────────────
def _tmp_save(uploaded, suffix=""):
    if uploaded is None:
        return None
    # Preserve the uploaded file's real extension (so an .xlsx upload isn't
    # saved as .csv). Fall back to any explicit suffix, then to the name's ext.
    ext = os.path.splitext(getattr(uploaded, "name", "") or "")[1]
    suffix = ext or suffix or ".bin"
    path = tempfile.NamedTemporaryFile(delete=False, suffix=suffix).name
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    return path


def _download_btn(path: str, label: str = "⬇ Download Excel"):
    with open(path, "rb") as f:
        data = f.read()
    st.download_button(
        label=label,
        data=data,
        file_name=os.path.basename(path),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=path,
    )


def _result_panel(rep):
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Period",     rep.month_label)
    m2.metric("Rate",       f"₹{rep.rate:.5f}")
    m3.metric("Accounts",   rep.n_accounts)
    m4.metric("Line rows",  rep.n_rows)
    st.success("Sheet generated successfully. Review it before sending to the client.")
    _download_btn(rep.out_path)


# ═══════════════════════════════════════════════════════════════════════════
# Main UI
# ═══════════════════════════════════════════════════════════════════════════
st.title("IDBI AWS Billing Tool")
st.caption("Applied Cloud Computing · internal tool · all processing is local")

tab_dbd, tab_dr, tab_itd = st.tabs([
    "📁 DBD Monthly Billing",
    "📋 DR / Estimate Sheet",
    "🏦 ITD Monthly Billing",
])


# ── TAB 1 : DBD billing ────────────────────────────────────────────────────
with tab_dbd:
    st.subheader("DBD Monthly Billing")
    st.markdown(
        "Accounts: DBD_Sandbox (528757802229), DBD-Network (417651893846), "
        "SB-Atlas (919437049652), Kaizen (562559071104), "
        "ImpactSure (294539078227), Atlas_Sandbox (689397855052)"
    )
    with st.expander("Pricing rules applied", expanded=False):
        st.markdown("""
- **BoM rows (orange):** Block SSD (slab), Snapshot separate @8%, Object Storage, 
  NAT Data, Static IP, ALB, Network Firewall @60%, KMS @99%, WAF, Net Transfer
- **EC2:** family-level substitution (c6/m6/r6 RHEL; **m5a & m6a → m5a BoM lines** for Windows; **Win r6 @8%**).
  Scaling by vCPU ratio, discount unchanged. No instance ever resolves to 0 vCPU.
- **EBS slab:** ≤128 → 128 GB; ≤256 → 256 GB; ≤512 → 512 GB; ≤1024 → 1024 GB; >1024 → 2048 GB
- **Direct Connect:** BoM leased-line ₹19,912.5/Gbps → (((19912.5/1000)*mbps)*qty)/rate, **75%**
- **Free tier:** zero-cost lines are excluded from the sheet
- **VPC:** IPv4 → Static IP BoM; TGW/Encryption/Endpoints → VPC Other @8%
- **Non-BoM @8%:** EKS, ECS, ECR, RDS, CloudWatch, Config, CloudTrail, OpenSearch, etc.
- CPU credits billed when present; compute rows tagged to their own account; Notes column only when notes exist.
""")

    col1, col2, col3 = st.columns(3)
    with col1: cur_f  = st.file_uploader("CUR CSV",        type=["csv","xlsx","xls"], key="dbd_cur")
    with col2: inv_f  = st.file_uploader("Invoice PDF",    type=["pdf"], key="dbd_inv")
    with col3: bom_f  = st.file_uploader("Section B BoM",  type=["xlsx"],key="dbd_bom")

    date_ov = st.text_input(
        "Rate date label (optional — auto-detected from invoice)",
        placeholder="e.g.  01st March 2026",
        key="dbd_date",
    )

    if st.button("Generate DBD Bill", type="primary",
                 disabled=not (cur_f and inv_f and bom_f)):
        with st.spinner("Processing…"):
            td = tempfile.mkdtemp()
            try:
                from idbi_billing import build_billing
                rep = build_billing(
                    cur_path     = _tmp_save(cur_f,  ".csv"),
                    invoice_path = _tmp_save(inv_f,  ".pdf"),
                    bom_path     = _tmp_save(bom_f,  ".xlsx"),
                    out_dir      = td,
                    group        = "dbd",
                    date_override= date_ov.strip() or None,
                )
                _result_panel(rep)
            except Exception as e:
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())


# ── TAB 2 : DR / Estimate ──────────────────────────────────────────────────
with tab_dr:
    st.subheader("DR / Cost Estimate Sheet")
    st.markdown("Upload an **AWS Pricing Calculator CSV export** + the Section B BoM.")

    with st.expander("Pricing rules applied", expanded=False):
        st.markdown("""
- **EC2:** family-level BoM substitution with vCPU-ratio scaling (discount % fixed).
  Windows r6a → no BoM equivalent → 8%.
- **EBS:** slab pricing (128/256/512/1024/2048 GB). EBS Snapshot separate @8%.
- **ALB, Network Firewall, WAF** → BoM lines with correct discounts.
- Everything else → calculator estimate @8%.
- Working Notes generated for every BoM substitution.
""")

    col1, col2 = st.columns(2)
    with col1: dr_csv  = st.file_uploader("Pricing Calculator CSV", type=["csv","xlsx","xls"],  key="dr_csv")
    with col2: dr_bom  = st.file_uploader("Section B BoM (xlsx)",   type=["xlsx"], key="dr_bom")

    col3, col4 = st.columns(2)
    with col3:
        dr_rate = st.number_input(
            "USD → INR conversion rate",
            min_value=50.0, max_value=120.0, value=91.07, step=0.01,
            key="dr_rate",
        )
    with col4:
        dr_date  = st.text_input("Rate date label", placeholder="e.g. 01st April 2026", key="dr_date")
        dr_title = st.text_input("Sheet / file title", value="DR Estimate", key="dr_title")

    if st.button("Generate DR Estimate", type="primary",
                 disabled=not (dr_csv and dr_bom)):
        with st.spinner("Processing…"):
            td = tempfile.mkdtemp()
            try:
                from idbi_billing.dr_estimate import build_estimate
                out_path = build_estimate(
                    csv_path   = _tmp_save(dr_csv, ".csv"),
                    bom_path   = _tmp_save(dr_bom, ".xlsx"),
                    out_dir    = td,
                    title      = dr_title or "DR Estimate",
                    rate       = float(dr_rate),
                    date_label = dr_date.strip() or f"Rs.{dr_rate}",
                )
                st.success("DR estimate sheet generated. Review all substitutions in Working Notes.")
                _download_btn(out_path)
            except Exception as e:
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())


# ── TAB 3 : ITD billing ────────────────────────────────────────────────────
with tab_itd:
    st.subheader("ITD Monthly Billing")
    st.markdown(
        "Accounts: IDBI Bank (122610479172), Network_Account (354929337302), "
        "Infra_Dev (647910641444), Infra_Prod (435200406974), "
        "Common_Services (789665322921), Audit (347887033541), "
        "Log_Archive (497783363028), Infra_Sandbox (403276262026), "
        "Infra_UAT (226202864192)"
    )

    with st.expander("Pricing rules applied", expanded=False):
        st.markdown("""
Same rules as DBD:
- BoM rows (orange) for Block SSD, Object Storage, NAT Data, Static IP, ALB, 
  Network Firewall, KMS, WAF, Net Transfer.
- EC2 family-level substitution with vCPU-ratio scaling.
- Non-BoM @8%: EKS, ECS, ECR, RDS, CloudWatch, Config, CloudTrail, etc.
- EBS slab pricing; Snapshot separate @8%.
""")

    col1, col2, col3 = st.columns(3)
    with col1: itd_cur = st.file_uploader("CUR CSV",       type=["csv","xlsx","xls"],  key="itd_cur")
    with col2: itd_inv = st.file_uploader("Invoice PDF",   type=["pdf"],  key="itd_inv")
    with col3: itd_bom = st.file_uploader("Section B BoM", type=["xlsx"], key="itd_bom")

    itd_date = st.text_input(
        "Rate date label (optional)",
        placeholder="e.g.  01st March 2026",
        key="itd_date",
    )

    if st.button("Generate ITD Bill", type="primary",
                 disabled=not (itd_cur and itd_inv and itd_bom)):
        with st.spinner("Processing…"):
            td = tempfile.mkdtemp()
            try:
                from idbi_billing import build_billing
                rep = build_billing(
                    cur_path     = _tmp_save(itd_cur, ".csv"),
                    invoice_path = _tmp_save(itd_inv, ".pdf"),
                    bom_path     = _tmp_save(itd_bom, ".xlsx"),
                    out_dir      = td,
                    group        = "itd",
                    date_override= itd_date.strip() or None,
                )
                _result_panel(rep)
            except Exception as e:
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())

# ── Footer ─────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "⚠️  This tool assists your review — always open the generated sheet and "
    "verify numbers before sending to the client."
)
