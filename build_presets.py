import os
import json

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
manifest_path = os.path.join(DATA_DIR, "manifest.json")

def main():
    if not os.path.exists(manifest_path):
        print("Manifest not found.")
        return

    with open(manifest_path, "r", encoding="utf-8") as f:
        manifest = json.load(f)

    all_data = {}
    for ticker, info in manifest.items():
        filepath = os.path.join(DATA_DIR, info["filename"])
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8") as f:
                all_data[ticker] = json.load(f)
            print(f"Loaded {ticker}")

    out_file = os.path.join(os.path.dirname(__file__), "data_bundle.js")
    with open(out_file, "w", encoding="utf-8") as f:
        f.write("// Auto-generated data bundle for offline & file:// execution\n")
        f.write("window.DCA_PRESETS = ")
        json.dump(all_data, f, ensure_ascii=False)
        f.write(";\n")
    print(f"Generated data_bundle.js (size: {os.path.getsize(out_file):,} bytes)")

if __name__ == "__main__":
    main()
