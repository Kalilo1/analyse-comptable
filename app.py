import re
import io
import json
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Analyse Comptable",
    page_icon="📊",
    layout="wide",
)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR — MENU PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════
# ── Menu dans la page ─────────────────────────────────────────────────────────
if "menu" not in st.session_state:
    st.session_state.menu = "🏠 Accueil"

st.markdown("## 📊 Analyse Comptable post-migration")
m1, m2, m3 = st.columns(3)
with m1:
    if st.button("🏠  Accueil", use_container_width=True,
                 type="primary" if st.session_state.menu == "🏠 Accueil" else "secondary"):
        st.session_state.menu = "🏠 Accueil"
        st.rerun()
with m2:
    if st.button("📒  Grand Livre", use_container_width=True,
                 type="primary" if st.session_state.menu == "📒 Grand Livre" else "secondary"):
        st.session_state.menu = "📒 Grand Livre"
        st.rerun()
with m3:
    if st.button("⚖️  Balance Auxiliaire", use_container_width=True,
                 type="primary" if st.session_state.menu == "⚖️ Balance Auxiliaire" else "secondary"):
        st.session_state.menu = "⚖️ Balance Auxiliaire"
        st.rerun()
st.markdown("---")
menu = st.session_state.menu


# ══════════════════════════════════════════════════════════════════════════════
# ACCUEIL
# ══════════════════════════════════════════════════════════════════════════════
if menu == "🏠 Accueil":
    st.title("📊 Analyse Comptable Fournisseurs")
    st.markdown("Sélectionnez un module dans le menu à gauche.")
    st.markdown("---")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
### 📒 Grand Livre
Comparaison de deux fichiers **Grand Livre Fournisseurs** au format TXT
(colonnes fixes, une ligne par transaction).

**Fonctionnalités :**
- Comparaison agrégée Débit / Crédit par fournisseur
- Détection des documents manquants par fournisseur
- Export Excel multi-onglets
        """)
        if st.button("Ouvrir Grand Livre →", use_container_width=True):
            st.session_state.menu = "📒 Grand Livre"
            st.rerun()

    with c2:
        st.markdown("""
### ⚖️ Balance Auxiliaire
Comparaison de deux fichiers **Balance Auxiliaire Fournisseurs** au format TXT
(2 lignes par fournisseur séparées par `|`).

**Fonctionnalités :**
- Balance antérieure, mouvements, balance finale, solde
- Fournisseurs communs, manquants, écarts
- Labels personnalisables (Fichier A / Fichier B)
- Export Excel multi-onglets
        """)
        if st.button("Ouvrir Balance Auxiliaire →", use_container_width=True):
            st.session_state.menu = "⚖️ Balance Auxiliaire"
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 1 — GRAND LIVRE
# ══════════════════════════════════════════════════════════════════════════════
elif menu == "📒 Grand Livre":

    st.title("📒 Comparaison Grand Livre Fournisseurs")

    # ── Parser ─────────────────────────────────────────────────────────────────
    @st.cache_data
    def parse_grand_livre(file_bytes: bytes, label: str = "fichier") -> pd.DataFrame:
        rows = []
        current_code = current_name = ""
        for line in file_bytes.decode("utf-8", errors="ignore").splitlines():
            line = line.replace('\r', '')
            m = re.match(r'^(F\d+)\s+(.*)', line)
            if m:
                current_code = m.group(1).strip()
                current_name = m.group(2).strip()
                continue
            if not re.match(r'^\d{2}/\d{2}/\d{2}', line):
                continue
            if re.search(r'(Tot du|Cumuls au|cumuls au)', line):
                continue
            if len(line) < 50:
                continue
            solde  = line[-14:].strip()
            credit = line[-28:-14].strip()
            debit  = line[-42:-28].strip()
            rest   = line[:-42].strip()
            parts  = rest.split()
            if len(parts) < 4:
                continue
            rows.append({
                "Fournisseur": current_code,
                "Nom":         current_name,
                "Date":        parts[0],
                "Document":    parts[1],
                "Type":        parts[2],
                "Reference":   parts[3],
                "Debit":       debit,
                "Credit":      credit,
                "Solde":       solde,
            })
        if not rows:
            st.warning(f"⚠️ **{label}** : aucune transaction détectée.")
            return pd.DataFrame(columns=["Fournisseur", "Nom", "Date", "Document",
                                         "Type", "Reference", "Debit", "Credit", "Solde"])
        df = pd.DataFrame(rows)
        for col in ["Debit", "Credit", "Solde"]:
            df[col] = (df[col].astype(str)
                       .str.replace(r'\s+', '', regex=True)
                       .str.replace(',', '.', regex=False))
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df

    # ── Missing documents ──────────────────────────────────────────────────────
    def gl_compute_missing(df1, df2, la, lb):
        all_suppliers = (pd.concat([df1[["Fournisseur", "Nom"]], df2[["Fournisseur", "Nom"]]])
                         .drop_duplicates("Fournisseur").sort_values("Fournisseur"))
        records = []
        for _, sup_row in all_suppliers.iterrows():
            code = sup_row["Fournisseur"]
            nom  = sup_row["Nom"]
            docs1 = set(df1.loc[df1["Fournisseur"] == code, "Document"])
            docs2 = set(df2.loc[df2["Fournisseur"] == code, "Document"])
            for doc in sorted(docs1 - docs2):
                r = df1[(df1["Fournisseur"] == code) & (df1["Document"] == doc)].iloc[0]
                records.append({"Fournisseur": code, "Nom": nom, "Document": doc,
                                "Présent dans": la, "Absent dans": lb,
                                "Date": r["Date"], "Type": r["Type"],
                                "Reference": r["Reference"], "Debit": r["Debit"], "Credit": r["Credit"]})
            for doc in sorted(docs2 - docs1):
                r = df2[(df2["Fournisseur"] == code) & (df2["Document"] == doc)].iloc[0]
                records.append({"Fournisseur": code, "Nom": nom, "Document": doc,
                                "Présent dans": lb, "Absent dans": la,
                                "Date": r["Date"], "Type": r["Type"],
                                "Reference": r["Reference"], "Debit": r["Debit"], "Credit": r["Credit"]})
        if not records:
            return pd.DataFrame()
        return pd.DataFrame(records).sort_values(["Fournisseur", "Présent dans", "Document"])

    # ── Excel ──────────────────────────────────────────────────────────────────
    def gl_build_excel(df1, df2, comp, missing, la, lb):
        wb = Workbook()
        thin = Side(style='thin', color='CCCCCC')
        brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

        def style_sheet(ws, df, title, tab_color):
            ws.title = title[:31]
            ws.sheet_properties.tabColor = tab_color
            hfill = PatternFill("solid", fgColor="1F4E79")
            hfont = Font(color="FFFFFF", bold=True, size=11)
            for ci, h in enumerate(df.columns, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill, c.font = hfill, hfont
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = brd
            for ri, row in enumerate(df.itertuples(index=False), 2):
                fill = PatternFill("solid", fgColor="EBF2FA" if ri % 2 == 0 else "FFFFFF")
                for ci, val in enumerate(row, 1):
                    c = ws.cell(row=ri, column=ci, value=val)
                    c.border = brd
                    c.fill   = fill
                    if isinstance(val, float):
                        c.number_format = '#,##0.000'
                        c.alignment = Alignment(horizontal='right')
                    else:
                        c.alignment = Alignment(horizontal='left')
            for ci, col in enumerate(df.columns, 1):
                w = max(len(str(col)), df[col].astype(str).str.len().max())
                ws.column_dimensions[get_column_letter(ci)].width = min(w + 4, 35)
            ws.row_dimensions[1].height = 20
            ws.freeze_panes = 'A2'

        ws1 = wb.active
        style_sheet(ws1, comp, "Comparaison", "1F4E79")
        red_f   = PatternFill("solid", fgColor="FFCCCC")
        green_f = PatternFill("solid", fgColor="CCFFCC")
        for ri in range(2, len(comp) + 2):
            for ci in [7, 8]:
                c = ws1.cell(row=ri, column=ci)
                if c.value and c.value != 0:
                    c.fill = red_f if c.value < 0 else green_f

        if not missing.empty:
            ws_m = wb.create_sheet()
            ws_m.title = "Documents manquants"
            ws_m.sheet_properties.tabColor = "C00000"
            sup_fill  = PatternFill("solid", fgColor="D9E1F2")
            sup_font  = Font(bold=True, size=11)
            red_fill  = PatternFill("solid", fgColor="FFE0E0")
            blue_fill = PatternFill("solid", fgColor="E0EEFF")
            cols = ["Document", "Présent dans", "Absent dans", "Date", "Type", "Reference", "Debit", "Credit"]
            ws_m.cell(row=1, column=1, value="Fournisseur / Nom").fill = PatternFill("solid", fgColor="1F4E79")
            ws_m.cell(row=1, column=1).font = Font(color="FFFFFF", bold=True, size=11)
            ws_m.cell(row=1, column=1).border = brd
            for ci, h in enumerate(cols, 2):
                c = ws_m.cell(row=1, column=ci, value=h)
                c.fill = PatternFill("solid", fgColor="1F4E79")
                c.font = Font(color="FFFFFF", bold=True, size=11)
                c.alignment = Alignment(horizontal='center')
                c.border = brd
            ws_m.row_dimensions[1].height = 20
            cur = 2
            for sup_code, grp in missing.groupby("Fournisseur", sort=True):
                ws_m.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=len(cols)+1)
                sc = ws_m.cell(row=cur, column=1, value=f"  {sup_code}  –  {grp['Nom'].iloc[0]}")
                sc.fill, sc.font = sup_fill, sup_font
                sc.alignment = Alignment(horizontal='left', vertical='center')
                sc.border = brd
                ws_m.row_dimensions[cur].height = 18
                cur += 1
                for _, drow in grp.iterrows():
                    rf = red_fill if drow["Absent dans"] == lb else blue_fill
                    ws_m.cell(row=cur, column=1, value="").fill = rf
                    for ci, col in enumerate(cols, 2):
                        val = drow[col]
                        c = ws_m.cell(row=cur, column=ci, value=val)
                        c.fill, c.border = rf, brd
                        if isinstance(val, float):
                            c.number_format = '#,##0.000'
                            c.alignment = Alignment(horizontal='right')
                        else:
                            c.alignment = Alignment(horizontal='left')
                    cur += 1
                cur += 1
            ws_m.column_dimensions["A"].width = 30
            for ci, col in enumerate(cols, 2):
                w = max(len(col), missing[col].astype(str).str.len().max() if col in missing.columns else 10)
                ws_m.column_dimensions[get_column_letter(ci)].width = min(w + 4, 35)
            ws_m.freeze_panes = 'A2'
            lr = cur + 1
            ws_m.cell(row=lr,   column=1, value="Légende :").font = Font(bold=True)
            ws_m.cell(row=lr+1, column=1, value=f"  Absent dans {lb}").fill = red_fill
            ws_m.cell(row=lr+2, column=1, value=f"  Absent dans {la}").fill = blue_fill

        ws3 = wb.create_sheet()
        style_sheet(ws3, df1, f"Détail {la}", "2E75B6")
        ws4 = wb.create_sheet()
        style_sheet(ws4, df2, f"Détail {lb}", "70AD47")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # ── Upload ─────────────────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    with c1:
        f1 = st.file_uploader("📂 Fichier A", type=["txt"], key="gl_f1")
        LA = st.text_input("Nom du fichier A", value="Fichier A", key="gl_la")
    with c2:
        f2 = st.file_uploader("📂 Fichier B", type=["txt"], key="gl_f2")
        LB = st.text_input("Nom du fichier B", value="Fichier B", key="gl_lb")

    if f1 and f2:
        with st.spinner("Parsing en cours..."):
            df1 = parse_grand_livre(f1.read(), LA)
            df2 = parse_grand_livre(f2.read(), LB)

        if df1.empty or df2.empty:
            st.error("Impossible de parser un ou plusieurs fichiers.")
            st.stop()

        sa, sb = f"_{LA}", f"_{LB}"
        agg1 = df1.groupby(["Fournisseur", "Nom"])[["Debit", "Credit"]].sum().reset_index()
        agg2 = df2.groupby(["Fournisseur", "Nom"])[["Debit", "Credit"]].sum().reset_index()
        comp = pd.merge(agg1, agg2, on="Fournisseur", how="outer", suffixes=(sa, sb))
        comp["Nom"] = comp[f"Nom{sa}"].fillna(comp[f"Nom{sb}"])
        comp = comp.drop(columns=[f"Nom{sa}", f"Nom{sb}"])
        comp = comp[["Fournisseur", "Nom",
                     f"Debit{sa}", f"Credit{sa}",
                     f"Debit{sb}", f"Credit{sb}"]].fillna(0)
        comp["Ecart_Debit"]  = comp[f"Debit{sb}"]  - comp[f"Debit{sa}"]
        comp["Ecart_Credit"] = comp[f"Credit{sb}"] - comp[f"Credit{sa}"]

        missing = gl_compute_missing(df1, df2, LA, LB)

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric(f"Fournisseurs {LA}", df1['Fournisseur'].nunique())
        k2.metric(f"Fournisseurs {LB}", df2['Fournisseur'].nunique())
        k3.metric(f"Lignes {LA}", len(df1))
        k4.metric(f"Lignes {LB}", len(df2))
        k5.metric("Documents manquants", len(missing) if not missing.empty else 0,
                  delta="⚠️" if not missing.empty else None, delta_color="inverse")

        tab1, tab2, tab3, tab4 = st.tabs([
            "📊 Comparaison agrégée",
            "🔍 Documents manquants",
            f"📄 Détail {LA}",
            f"📄 Détail {LB}",
        ])

        with tab1:
            fmt = {c: "{:,.3f}" for c in comp.columns if comp[c].dtype == float}
            st.dataframe(comp.style.format(fmt), use_container_width=True)

        with tab2:
            if missing.empty:
                st.success("✅ Aucun document manquant.")
            else:
                st.info(f"**{len(missing)} document(s) manquant(s)** sur "
                        f"**{missing['Fournisseur'].nunique()} fournisseur(s)**")
                for sup_code, grp in missing.groupby("Fournisseur", sort=True):
                    nom = grp["Nom"].iloc[0]
                    only_a = grp[grp["Absent dans"] == LB]
                    only_b = grp[grp["Absent dans"] == LA]
                    label = (f"**{sup_code}** – {nom}  "
                             + (f"🔴 {len(only_a)} absent(s) dans {LB}  " if not only_a.empty else "")
                             + (f"🔵 {len(only_b)} absent(s) dans {LA}"   if not only_b.empty else ""))
                    with st.expander(label, expanded=False):
                        if not only_a.empty:
                            st.markdown(f"🔴 **Présents dans {LA} — Absents dans {LB}**")
                            st.dataframe(only_a[["Document","Date","Type","Reference","Debit","Credit"]]
                                         .reset_index(drop=True), use_container_width=True)
                        if not only_b.empty:
                            st.markdown(f"🔵 **Présents dans {LB} — Absents dans {LA}**")
                            st.dataframe(only_b[["Document","Date","Type","Reference","Debit","Credit"]]
                                         .reset_index(drop=True), use_container_width=True)

        with tab3:
            st.dataframe(df1.style.format(
                {c: "{:,.3f}" for c in df1.columns if df1[c].dtype == float}
            ), use_container_width=True)

        with tab4:
            st.dataframe(df2.style.format(
                {c: "{:,.3f}" for c in df2.columns if df2[c].dtype == float}
            ), use_container_width=True)

        st.divider()
        excel_buf = gl_build_excel(df1, df2, comp, missing, LA, LB)
        st.download_button(
            label="📥 Télécharger Excel",
            data=excel_buf,
            file_name="comparaison_grand_livre.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ══════════════════════════════════════════════════════════════════════════════
# MODULE 2 — BALANCE AUXILIAIRE
# ══════════════════════════════════════════════════════════════════════════════
elif menu == "⚖️ Balance Auxiliaire":

    st.title("⚖️ Comparaison Balance Auxiliaire Fournisseurs")

    # ── Parser ─────────────────────────────────────────────────────────────────
    @st.cache_data
    def parse_balance(file_bytes: bytes, label: str = "fichier") -> pd.DataFrame:
        def to_float(s: str) -> float:
            s = s.strip().replace(' ', '').replace(',', '.')
            try:
                return float(s)
            except ValueError:
                return 0.0

        def extract_trailing_number(seg: str) -> float:
            m = re.search(r'([\d.]+)\s*$', seg.strip())
            return to_float(m.group(1)) if m else 0.0

        def clean_name(raw: str) -> str:
            return re.sub(r'\s+[\d,.]+\s*$', '', raw).strip()

        rows = []
        lines = file_bytes.decode("utf-8", errors="ignore").splitlines()
        lines = [l.replace('\r', '') for l in lines]
        i = 0
        while i < len(lines):
            line = lines[i]
            if re.match(r'^[Ff]\d+', line):
                fline = line
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    j += 1
                nline = lines[j] if j < len(lines) else ''
                fparts = fline.split('|')
                nparts = nline.split('|')
                code_m = re.match(r'^([Ff]\d+)', fparts[0])
                code   = code_m.group(1).upper() if code_m else ''
                name   = clean_name(nparts[0]) if nparts else ''
                rows.append({
                    "Fournisseur":   code,
                    "Nom":           name,
                    "BalAnt_Debit":  extract_trailing_number(fparts[0]),
                    "BalAnt_Credit": extract_trailing_number(nparts[0]) if nparts else 0.0,
                    "Mvt_Debit":     to_float(fparts[1]) if len(fparts) > 1 else 0.0,
                    "Mvt_Credit":    to_float(nparts[1]) if len(nparts) > 1 else 0.0,
                    "Bal_Debit":     to_float(fparts[2]) if len(fparts) > 2 else 0.0,
                    "Bal_Credit":    to_float(nparts[2]) if len(nparts) > 2 else 0.0,
                    "Solde_Debit":   to_float(fparts[3]) if len(fparts) > 3 else 0.0,
                    "Solde_Credit":  to_float(nparts[3]) if len(nparts) > 3 else 0.0,
                })
                i = j + 1
            else:
                i += 1

        if not rows:
            st.warning(f"⚠️ **{label}** : aucun fournisseur détecté.")
            return pd.DataFrame(columns=["Fournisseur", "Nom",
                                         "BalAnt_Debit", "BalAnt_Credit",
                                         "Mvt_Debit", "Mvt_Credit",
                                         "Bal_Debit", "Bal_Credit",
                                         "Solde_Debit", "Solde_Credit"])
        return pd.DataFrame(rows)

    # ── Communs ────────────────────────────────────────────────────────────────
    def ba_compute_common(df1, df2, la, lb):
        codes = set(df1["Fournisseur"]) & set(df2["Fournisseur"])
        sa, sb = f"_{la}", f"_{lb}"
        merged = pd.merge(df1[df1["Fournisseur"].isin(codes)],
                          df2[df2["Fournisseur"].isin(codes)],
                          on="Fournisseur", suffixes=(sa, sb))
        merged["Nom"] = merged[f"Nom{sa}"].fillna(merged[f"Nom{sb}"])
        merged = merged.drop(columns=[f"Nom{sa}", f"Nom{sb}"])
        cols = ["Fournisseur", "Nom"] + [c for c in merged.columns if c not in ("Fournisseur", "Nom")]
        return merged[cols].sort_values("Fournisseur").reset_index(drop=True)

    # ── Manquants ──────────────────────────────────────────────────────────────
    def ba_compute_missing(df1, df2, la, lb):
        codes1, codes2 = set(df1["Fournisseur"]), set(df2["Fournisseur"])
        records = []
        for code in sorted(codes1 - codes2):
            r = df1[df1["Fournisseur"] == code].iloc[0]
            records.append({"Fournisseur": code, "Nom": r["Nom"],
                            "Présent dans": la, "Absent dans": lb,
                            "BalAnt_Debit": r["BalAnt_Debit"], "BalAnt_Credit": r["BalAnt_Credit"],
                            "Mvt_Debit": r["Mvt_Debit"], "Mvt_Credit": r["Mvt_Credit"],
                            "Bal_Debit": r["Bal_Debit"], "Bal_Credit": r["Bal_Credit"],
                            "Solde_Debit": r["Solde_Debit"], "Solde_Credit": r["Solde_Credit"]})
        for code in sorted(codes2 - codes1):
            r = df2[df2["Fournisseur"] == code].iloc[0]
            records.append({"Fournisseur": code, "Nom": r["Nom"],
                            "Présent dans": lb, "Absent dans": la,
                            "BalAnt_Debit": r["BalAnt_Debit"], "BalAnt_Credit": r["BalAnt_Credit"],
                            "Mvt_Debit": r["Mvt_Debit"], "Mvt_Credit": r["Mvt_Credit"],
                            "Bal_Debit": r["Bal_Debit"], "Bal_Credit": r["Bal_Credit"],
                            "Solde_Debit": r["Solde_Debit"], "Solde_Credit": r["Solde_Credit"]})
        if not records:
            return pd.DataFrame()
        return pd.DataFrame(records).sort_values(["Absent dans", "Fournisseur"])

    # ── Excel ──────────────────────────────────────────────────────────────────
    def ba_build_excel(df1, df2, comp, missing, common, la, lb):
        wb   = Workbook()
        thin = Side(style='thin', color='CCCCCC')
        brd  = Border(left=thin, right=thin, top=thin, bottom=thin)

        def style_sheet(ws, df, title, tab_color):
            ws.title = title[:31]
            ws.sheet_properties.tabColor = tab_color
            hfill = PatternFill("solid", fgColor="1F4E79")
            hfont = Font(color="FFFFFF", bold=True, size=11)
            for ci, h in enumerate(df.columns, 1):
                c = ws.cell(row=1, column=ci, value=h)
                c.fill, c.font = hfill, hfont
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = brd
            for ri, row in enumerate(df.itertuples(index=False), 2):
                fill = PatternFill("solid", fgColor="EBF2FA" if ri % 2 == 0 else "FFFFFF")
                for ci, val in enumerate(row, 1):
                    c = ws.cell(row=ri, column=ci, value=val)
                    c.border = brd
                    c.fill   = fill
                    if isinstance(val, float):
                        c.number_format = '#,##0.000'
                        c.alignment = Alignment(horizontal='right')
                    else:
                        c.alignment = Alignment(horizontal='left')
            for ci, col in enumerate(df.columns, 1):
                w = max(len(str(col)), df[col].astype(str).str.len().max())
                ws.column_dimensions[get_column_letter(ci)].width = min(w + 4, 35)
            ws.row_dimensions[1].height = 20
            ws.freeze_panes = 'A2'

        ws1 = wb.active
        style_sheet(ws1, comp, "Comparaison", "1F4E79")
        red_f   = PatternFill("solid", fgColor="FFCCCC")
        green_f = PatternFill("solid", fgColor="CCFFCC")
        ecart_start = comp.columns.tolist().index("Ecart_BalAnt_Debit") + 1
        for ri in range(2, len(comp) + 2):
            for ci in range(ecart_start, ecart_start + 8):
                c = ws1.cell(row=ri, column=ci)
                if c.value and abs(c.value) > 0.001:
                    c.fill = red_f if c.value < 0 else green_f

        ws_com = wb.create_sheet()
        style_sheet(ws_com, common, "Fournisseurs communs", "375623")

        if not missing.empty:
            ws_m = wb.create_sheet()
            ws_m.title = "Fournisseurs manquants"
            ws_m.sheet_properties.tabColor = "C00000"
            hfill    = PatternFill("solid", fgColor="1F4E79")
            hfont    = Font(color="FFFFFF", bold=True, size=11)
            red_fill = PatternFill("solid", fgColor="FFE0E0")
            blu_fill = PatternFill("solid", fgColor="E0EEFF")
            for ci, h in enumerate(missing.columns, 1):
                c = ws_m.cell(row=1, column=ci, value=h)
                c.fill, c.font = hfill, hfont
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = brd
            ws_m.row_dimensions[1].height = 20
            for ri, row in enumerate(missing.itertuples(index=False), 2):
                absent = row[3]
                rf = red_fill if absent == lb else blu_fill
                for ci, val in enumerate(row, 1):
                    c = ws_m.cell(row=ri, column=ci, value=val)
                    c.fill, c.border = rf, brd
                    if isinstance(val, float):
                        c.number_format = '#,##0.000'
                        c.alignment = Alignment(horizontal='right')
                    else:
                        c.alignment = Alignment(horizontal='left')
            for ci, col in enumerate(missing.columns, 1):
                w = max(len(str(col)), missing[col].astype(str).str.len().max())
                ws_m.column_dimensions[get_column_letter(ci)].width = min(w + 4, 35)
            ws_m.freeze_panes = 'A2'
            lr = len(missing) + 3
            ws_m.cell(row=lr,   column=1, value="Légende :").font = Font(bold=True)
            ws_m.cell(row=lr+1, column=1, value=f"  Absent dans {lb}").fill = red_fill
            ws_m.cell(row=lr+2, column=1, value=f"  Absent dans {la}").fill = blu_fill

        ws4 = wb.create_sheet()
        style_sheet(ws4, df1, f"Données {la}", "2E75B6")
        ws5 = wb.create_sheet()
        style_sheet(ws5, df2, f"Données {lb}", "70AD47")

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return buf

    # ── Upload ─────────────────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    with c1:
        f1 = st.file_uploader("📂 Fichier A", type=["txt"], key="ba_f1")
        LA = st.text_input("Nom du fichier A", value="Fichier A", key="ba_la")
    with c2:
        f2 = st.file_uploader("📂 Fichier B", type=["txt"], key="ba_f2")
        LB = st.text_input("Nom du fichier B", value="Fichier B", key="ba_lb")

    if f1 and f2:
        with st.spinner("Parsing en cours..."):
            df1 = parse_balance(f1.read(), LA)
            df2 = parse_balance(f2.read(), LB)

        if df1.empty or df2.empty:
            st.error("Impossible de parser un ou plusieurs fichiers.")
            st.stop()

        sa, sb = f"_{LA}", f"_{LB}"
        comp = pd.merge(df1, df2, on="Fournisseur", how="outer", suffixes=(sa, sb))
        comp["Nom"] = comp[f"Nom{sa}"].fillna(comp[f"Nom{sb}"])
        comp = comp.drop(columns=[f"Nom{sa}", f"Nom{sb}"]).fillna(0)
        base_cols = ["Fournisseur", "Nom",
                     f"BalAnt_Debit{sa}", f"BalAnt_Credit{sa}",
                     f"BalAnt_Debit{sb}", f"BalAnt_Credit{sb}",
                     f"Mvt_Debit{sa}",    f"Mvt_Credit{sa}",
                     f"Mvt_Debit{sb}",    f"Mvt_Credit{sb}",
                     f"Bal_Debit{sa}",    f"Bal_Credit{sa}",
                     f"Bal_Debit{sb}",    f"Bal_Credit{sb}",
                     f"Solde_Debit{sa}",  f"Solde_Credit{sa}",
                     f"Solde_Debit{sb}",  f"Solde_Credit{sb}"]
        comp = comp[base_cols]
        comp["Ecart_BalAnt_Debit"]  = comp[f"BalAnt_Debit{sb}"]  - comp[f"BalAnt_Debit{sa}"]
        comp["Ecart_BalAnt_Credit"] = comp[f"BalAnt_Credit{sb}"] - comp[f"BalAnt_Credit{sa}"]
        comp["Ecart_Mvt_Debit"]     = comp[f"Mvt_Debit{sb}"]     - comp[f"Mvt_Debit{sa}"]
        comp["Ecart_Mvt_Credit"]    = comp[f"Mvt_Credit{sb}"]    - comp[f"Mvt_Credit{sa}"]
        comp["Ecart_Bal_Debit"]     = comp[f"Bal_Debit{sb}"]     - comp[f"Bal_Debit{sa}"]
        comp["Ecart_Bal_Credit"]    = comp[f"Bal_Credit{sb}"]    - comp[f"Bal_Credit{sa}"]
        comp["Ecart_Solde_Debit"]   = comp[f"Solde_Debit{sb}"]   - comp[f"Solde_Debit{sa}"]
        comp["Ecart_Solde_Credit"]  = comp[f"Solde_Credit{sb}"]  - comp[f"Solde_Credit{sa}"]

        missing = ba_compute_missing(df1, df2, LA, LB)
        common  = ba_compute_common(df1, df2, LA, LB)

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric(f"Fournisseurs {LA}", len(df1))
        k2.metric(f"Fournisseurs {LB}", len(df2))
        k3.metric("Communs", len(set(df1["Fournisseur"]) & set(df2["Fournisseur"])))
        k4.metric(f"Uniq. {LA}", len(set(df1["Fournisseur"]) - set(df2["Fournisseur"])))
        k5.metric(f"Uniq. {LB}", len(set(df2["Fournisseur"]) - set(df1["Fournisseur"])))
        nb_ecart = len(comp[
            (comp["Ecart_BalAnt_Debit"].abs()  > 0.01) |
            (comp["Ecart_BalAnt_Credit"].abs() > 0.01) |
            (comp["Ecart_Mvt_Debit"].abs()     > 0.01) |
            (comp["Ecart_Mvt_Credit"].abs()    > 0.01)
        ])
        k6.metric("Avec écart", nb_ecart,
                  delta="⚠️" if nb_ecart > 0 else None, delta_color="inverse")

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Comparaison",
            "🤝 Fournisseurs communs",
            "🔍 Fournisseurs manquants",
            f"📄 Données {LA}",
            f"📄 Données {LB}",
        ])

        fmt = {c: "{:,.3f}" for c in comp.columns if comp[c].dtype == float}

        with tab1:
            fc1, fc2 = st.columns([2, 1])
            with fc2:
                only_ecart = st.toggle("Afficher uniquement les écarts", value=False, key="ba_toggle")
            display = comp.copy()
            if only_ecart:
                display = display[
                    (display["Ecart_BalAnt_Debit"].abs()  > 0.01) |
                    (display["Ecart_BalAnt_Credit"].abs() > 0.01) |
                    (display["Ecart_Mvt_Debit"].abs()     > 0.01) |
                    (display["Ecart_Mvt_Credit"].abs()    > 0.01)
                ]
                st.caption(f"{len(display)} fournisseur(s) avec écart")
            st.dataframe(display.style.format(fmt), use_container_width=True)

        with tab2:
            st.caption(f"{len(common)} fournisseurs présents dans les deux fichiers")
            fmt_c = {c: "{:,.3f}" for c in common.columns if common[c].dtype == float}
            st.dataframe(common.style.format(fmt_c), use_container_width=True)

        with tab3:
            if missing.empty:
                st.success("✅ Tous les fournisseurs sont présents dans les deux fichiers.")
            else:
                only_a = missing[missing["Absent dans"] == LB]
                only_b = missing[missing["Absent dans"] == LA]
                st.info(
                    f"**{len(missing)} fournisseur(s) manquant(s)** — "
                    f"🔴 {len(only_a)} absent(s) dans {LB} · "
                    f"🔵 {len(only_b)} absent(s) dans {LA}"
                )
                if not only_a.empty:
                    st.markdown(f"#### 🔴 Présents dans {LA} — Absents dans {LB}")
                    st.dataframe(only_a.drop(columns=["Présent dans", "Absent dans"])
                                 .reset_index(drop=True)
                                 .style.format({c: "{:,.3f}" for c in only_a.columns
                                                if only_a[c].dtype == float}),
                                 use_container_width=True)
                if not only_b.empty:
                    st.markdown(f"#### 🔵 Présents dans {LB} — Absents dans {LA}")
                    st.dataframe(only_b.drop(columns=["Présent dans", "Absent dans"])
                                 .reset_index(drop=True)
                                 .style.format({c: "{:,.3f}" for c in only_b.columns
                                                if only_b[c].dtype == float}),
                                 use_container_width=True)

        with tab4:
            st.dataframe(df1.style.format(
                {c: "{:,.3f}" for c in df1.columns if df1[c].dtype == float}
            ), use_container_width=True)

        with tab5:
            st.dataframe(df2.style.format(
                {c: "{:,.3f}" for c in df2.columns if df2[c].dtype == float}
            ), use_container_width=True)

        st.divider()
        excel_buf = ba_build_excel(df1, df2, comp, missing, common, LA, LB)
        st.download_button(
            label="📥 Télécharger Excel (5 onglets)",
            data=excel_buf,
            file_name="comparaison_balance_auxiliaire.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
