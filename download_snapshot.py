import urllib.request
import os

URL = "https://github.com/nixervo/Clairvoyant-Reps/raw/main/clan_2527.xlsx"
OUTPUT = "clan_2527.xlsx"

req = urllib.request.Request(URL, headers={"User-Agent": "download-snapshot/1.0"})
with urllib.request.urlopen(req) as resp:
    data = resp.read()

with open(OUTPUT, "wb") as f:
    f.write(data)

print(f"Downloaded {len(data)} bytes -> {OUTPUT}")
