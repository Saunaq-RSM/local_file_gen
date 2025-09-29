import streamlit as st
from processor import configure, process_and_fill

# Configure backend with secrets
configure(
    st.secrets["AZURE_API_KEY"],
    st.secrets["AZURE_API_ENDPOINT"]
)

st.set_page_config(page_title="TP Agent 1 Reinvoicing", layout="wide")
st.title("Transfer Pricing Document Filler")

uploaded = st.file_uploader(
    "Upload: OECD text (.txt), transcript (.docx), analysis (.pdf), variables (.xlsx), and template (.pptx/.docx).",
    type=["txt", "docx", "pdf", "xlsx", "pptx"],
    accept_multiple_files=True
)

file_map = {}
if uploaded:
    # First pass: direct mapping by filename hints
    txts, docxs, pdfs, xlsxs, pptxs = [], [], [], [], []
    for f in uploaded:
        n = f.name.lower()
        if n.endswith(".txt"):
            txts.append(f)
            if "oecd" in n or "guideline" in n:
                file_map.setdefault("guidelines", f)
        elif n.endswith(".pdf"):
            pdfs.append(f)
            if "analysis" in n:
                file_map.setdefault("pdf", f)
        elif n.endswith(".xlsx"):
            xlsxs.append(f)
            if "var" in n or "replace" in n:
                file_map.setdefault("excel", f)
        elif n.endswith(".pptx"):
            pptxs.append(f)
            if "template" in n:
                file_map.setdefault("template", f)
        elif n.endswith(".docx"):
            docxs.append(f)
            if "transcript" in n or "interview" in n:
                file_map.setdefault("transcript", f)
            if "template" in n:
                file_map.setdefault("template", f)

    # Fallbacks: pick first available for any missing role
    if "guidelines" not in file_map and txts:
        file_map["guidelines"] = txts[0]
    if "pdf" not in file_map and pdfs:
        file_map["pdf"] = pdfs[0]
    if "excel" not in file_map and xlsxs:
        file_map["excel"] = xlsxs[0]

    # Template preference: PPTX first, else DOCX
    if "template" not in file_map:
        if pptxs:
            file_map["template"] = pptxs[-1]
        elif docxs:
            candidate = docxs[-1]
            if "transcript" in candidate.name.lower() and len(docxs) > 1:
                candidate = [d for d in docxs if "transcript" not in d.name.lower()][-1]
            file_map["template"] = candidate

    # Transcript fallback: any DOCX that's not the template
    if "transcript" not in file_map and docxs:
        candidates = [d for d in docxs if d is not file_map.get("template")]
        if candidates:
            file_map["transcript"] = candidates[0]

# Ensure all expected keys exist (default to empty string)
for key in ["guidelines", "pdf", "excel", "template", "transcript"]:
    file_map.setdefault(key, "")

if uploaded:
    with st.expander("Detected files (auto-mapped)"):
        for k, v in file_map.items():
            if v:
                st.write(f"- **{k}** → {v.name}")
            else:
                st.write(f"- **{k}** → (none)")

    if st.button("Generate filled file"):
        with st.spinner("Processing..."):
            try:
                path = process_and_fill(file_map)
                st.success("Done—download below:")
                if path.endswith(".pptx"):
                    mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    dl_name = "filled.pptx"
                else:
                    mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    dl_name = "filled.docx"
                st.download_button(
                    f"Download {dl_name}",
                    open(path, "rb"),
                    file_name=dl_name,
                    mime=mime,
                )
            except Exception as e:
                st.error(f"Error: {e}")
