"""
Connect to a running Power BI Desktop instance via its Analysis Services port.

Requires: pip install pythonnet pyadomd
And: Power BI Desktop installed (DLL ships with it).
"""
import subprocess, re, clr


ADOMD_DLL = r"C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.PowerBI.AdomdClient.dll"


def find_port():
    """Return the localhost port Power BI Desktop's Analysis Services is listening on."""
    out = subprocess.check_output(
        ["powershell", "-NoProfile", "-Command",
         "Get-Process msmdsrv | ForEach-Object { "
         "  Get-NetTCPConnection -OwningProcess $_.Id -State Listen | "
         "  Where-Object { $_.LocalAddress -eq '127.0.0.1' } | "
         "  Select-Object -First 1 -ExpandProperty LocalPort }"],
        text=True)
    m = re.search(r"\d+", out)
    if not m:
        raise RuntimeError("Power BI Desktop not running, or no AS port found.")
    return int(m.group())


def connect(port=None):
    if port is None:
        port = find_port()
    clr.AddReference(ADOMD_DLL)
    # Real namespace is Microsoft.AnalysisServices.AdomdClient (not the DLL name)
    from Microsoft.AnalysisServices.AdomdClient import AdomdConnection
    conn = AdomdConnection(f"Data Source=localhost:{port}")
    conn.Open()
    return conn


def dax(conn, query):
    """Run a DAX query, return list of dicts."""
    cmd = conn.CreateCommand()
    cmd.CommandText = query
    rdr = cmd.ExecuteReader()
    cols = [rdr.GetName(i) for i in range(rdr.FieldCount)]
    rows = []
    while rdr.Read():
        rows.append({c: rdr[i] for i, c in enumerate(cols)})
    rdr.Close()
    return rows


if __name__ == "__main__":
    c = connect()
    print("Tables:")
    for r in dax(c, "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES"):
        print(" -", r.get("Name"))
    c.Close()
