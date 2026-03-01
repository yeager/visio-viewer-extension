(async function() {
  let pyodide = null;

  async function init() {
    pyodide = await loadPyodide({
      indexURL: "pyodide/",
      lockFileURL: "pyodide/pyodide-lock.json"
    });
    await pyodide.loadPackage("micropip");
    const micropip = pyodide.pyimport("micropip");
    // Install from local bundled wheels
    await micropip.install("pyodide/olefile-0.47-py2.py3-none-any.whl");
    await micropip.install("pyodide/libvisio_ng-0.6.0-py3-none-any.whl");
    return true;
  }

  async function convertFile(data) {
    const uint8 = new Uint8Array(data);
    pyodide.FS.writeFile("/tmp/input.vsdx", uint8);

    const result = await pyodide.runPythonAsync(`
import json, os
from libvisio_ng import convert, get_page_info

pages_info = get_page_info("/tmp/input.vsdx")
page_names = [p.get("name", f"Page {p['index']+1}") for p in pages_info]

svg_files = convert("/tmp/input.vsdx", output_dir="/tmp/svg_out")
svgs = []
for f in sorted(svg_files):
    with open(f, "r") as fh:
        svgs.append(fh.read())

# Cleanup
for f in svg_files:
    os.remove(f)
os.remove("/tmp/input.vsdx")

json.dumps({"svgs": svgs, "names": page_names})
`);
    return JSON.parse(result);
  }

  window.addEventListener("message", async (event) => {
    const { type, id, data } = event.data;
    try {
      if (type === "init") {
        await init();
        parent.postMessage({ type: "init-done", id }, "*");
      } else if (type === "convert") {
        const result = await convertFile(data);
        parent.postMessage({ type: "convert-done", id, result }, "*");
      }
    } catch (err) {
      parent.postMessage({ type: "error", id, error: err.message || String(err) }, "*");
    }
  });

  parent.postMessage({ type: "sandbox-ready" }, "*");
})();
