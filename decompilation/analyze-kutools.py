# analyze-kutools.py — Phase 3a groundwork: enumerate .NET metadata (TypeDef inventory
# + Word-interop MemberRef calls) from KuTools assemblies without ILSpy.
# Usage: python decompilation/analyze-kutools.py
import dnfile, os, collections

DLLS = [
    r"C:\Program Files (x86)\ExtendOffice\WordKutools\Kutools\Kutools.dll",
    r"C:\Program Files (x86)\ExtendOffice\WordKutools\Kutools4\Kutools.dll",
    r"C:\Program Files (x86)\ExtendOffice\WordKutools\Kutools\Utils.dll",
]
OUT = r"C:\Users\Jaja\AppData\Local\Temp\opencode"


def typename(typedef):
    ns = str(typedef.TypeNamespace or "")
    nm = str(typedef.TypeName or "")
    return (ns + "." + nm) if ns else nm


def main():
    all_typedefs = []
    word_calls = collections.defaultdict(set)  # (type namespace.name) -> {member names}
    for dll in DLLS:
        tag = os.path.basename(os.path.dirname(dll)) + "/" + os.path.basename(dll)
        try:
            pe = dnfile.dnPE(dll)
        except Exception as e:  # noqa: BLE001
            print(f"{tag}: PARSE FAILED ({e})")
            continue
        if pe.net is None:
            print(f"{tag}: no .NET metadata (native image)")
            continue
        tables = pe.net.mdtables
        typedefs = [typename(t) for t in tables.TypeDef]
        all_typedefs += [(tag, t) for t in typedefs]
        # map TypeRef/TypeDef 1-based row -> full name for MemberRefParent resolution
        typerefs = {}
        for i, tr in enumerate(tables.TypeRef, start=1):
            ns = str(tr.TypeNamespace or "")
            nm = str(tr.TypeName or "")
            typerefs[i] = (ns + "." + nm) if ns else nm
        typedef_names = [typename(t) for t in tables.TypeDef]
        n_calls = 0
        for mr in tables.MemberRef:
            try:
                parent_table = mr.Class.table
                parent_row = mr.Class.row_index
            except Exception:  # noqa: BLE001
                continue
            if parent_table == "TypeRef":
                full = typerefs.get(parent_row, "")
            elif parent_table == "TypeDef":
                full = typedef_names[parent_row - 1] if 0 < parent_row <= len(typedef_names) else ""
            else:
                continue
            if full.startswith("Microsoft.Office.Interop.Word") or full == "Word":
                word_calls[full].add(str(mr.Name))
                n_calls += 1
        n_methods = len(tables.MethodDef)
        print(f"{tag}: {len(typedefs)} types, {n_methods} methods, {n_calls} Word-interop MemberRefs")

    with open(os.path.join(OUT, "kutools-typedefs.txt"), "w", encoding="utf-8") as f:
        for tag, t in sorted(set(all_typedefs)):
            f.write(f"{tag} :: {t}\n")
    with open(os.path.join(OUT, "kutools-wordcalls.txt"), "w", encoding="utf-8") as f:
        for t in sorted(word_calls):
            f.write(f"{t}: {', '.join(sorted(word_calls[t]))}\n")
    print(f"typedefs total: {len(set(all_typedefs))}, word-call types: {len(word_calls)}")
    print("saved kutools-typedefs.txt + kutools-wordcalls.txt")


if __name__ == "__main__":
    main()
