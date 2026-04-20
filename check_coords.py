# -*- coding: utf-8 -*-
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

with open('map_data.json', encoding='utf-8') as f:
    data = json.load(f)

for sh in list(data.keys()):
    rows = data[sh]
    print(f'Sheet: {sh}, total rows: {len(rows)}')
    for r in rows[:2]:
        print(f'  lon_a={r["lon_a"]}, lat_a={r["lat_a"]}, lon_b={r["lon_b"]}, lat_b={r["lat_b"]}')
    # check for bad values
    bad = [r for r in rows if r['lon_a'] is not None and (r['lon_a'] > 180 or r['lon_a'] < -180)]
    print(f'  Bad lon_a count: {len(bad)}')
    if bad:
        print(f'  Example bad lon_a: {bad[0]["lon_a"]}')
    bad_lat = [r for r in rows if r['lat_a'] is not None and (r['lat_a'] > 90 or r['lat_a'] < -90)]
    print(f'  Bad lat_a count: {len(bad_lat)}')
