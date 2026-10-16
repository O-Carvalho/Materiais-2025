import glob, re, xlrd, xlwt

def processa(path):
    b = xlrd.open_workbook(path, formatting_info=True)
    s = b.sheet_by_index(0)
    if s.cell_value(0, 0) == "OS":
        return "pulado"
    os_num = int(re.fullmatch(r"OS (\d+)\.xls", path).group(1))
    estilos = {}
    def estilo(r, c):
        xf = b.xf_list[s.cell_xf_index(r, c)]
        fmt = b.format_map[xf.format_key].format_str
        bold = b.font_list[xf.font_index].bold
        chave = (fmt, bold)
        if chave not in estilos:
            st = xlwt.XFStyle()
            st.num_format_str = fmt
            st.font.bold = bool(bold)
            estilos[chave] = st
        return estilos[chave]
    wb = xlwt.Workbook()
    ws = wb.add_sheet(s.name)
    head = estilo(0, 0)
    ws.write(0, 0, "OS", head)
    ws.col(0).width = 2500
    for c in range(s.ncols):
        if c in s.colinfo_map:
            ws.col(c + 1).width = s.colinfo_map[c].width
    for r in range(s.nrows):
        if r > 0:
            ws.write(r, 0, os_num)
        for c in range(s.ncols):
            if s.cell_type(r, c) == xlrd.XL_CELL_EMPTY:
                continue
            ws.write(r, c + 1, s.cell_value(r, c), estilo(r, c))
    wb.save(path)
    return "ok"

if __name__ == "__main__":
    from collections import Counter
    print(Counter(processa(f) for f in sorted(glob.glob("OS *.xls"))))
