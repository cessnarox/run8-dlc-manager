#!/usr/bin/env python3
"""
run8dlc - Run8 Train Simulator DLC library manager
===================================================
Tracks what Run8 DLC you own vs. what's on the 3DTS store, migrates an
existing Installers/Transactions folder setup into a ledger, and captures
new purchases (downloads the EXE, screenshots the receipt page, records
the transaction, optionally launches the installer).

Commands:
  report     (default) diff owned vs store catalog -> console + report.html
  installed  list every product with its install status
  migrate    ingest existing Installers/ + Transactions/ into ledger.json
  add        record a new purchase from a receipt-page URL or a local EXE
  reinstall  launch the installer EXE for a product you own
  uninstall  quarantine a product's in-game folders (reversible)
  restore    put a quarantined product's folders back
  updater    launch Run8_Updater.exe
  refresh    re-scan the store for new products / update prices
  snapshot   dump a directory listing of your Run8 install for diagnostics

Stdlib only. Windows-first but runs anywhere for read-only commands.
GUI: run8dlc_gui.py wraps all of this in a window.
"""

import argparse
import difflib
import html
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.request
import urllib.parse
from datetime import datetime
from pathlib import Path

def _resolve_app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _resolve_data_dir():
    """Everything the app creates lives in a 'data' folder next to
    the program, like a normal install -- the app root stays clean
    (program + README). Setups from before v0.9.16 kept files loose
    next to the app; they are moved into 'data' automatically, once."""
    base = _resolve_app_dir()
    probe = base / ".r8dlc_write_test"
    try:
        probe.touch()
        probe.unlink()
    except OSError:
        alt = Path(os.environ.get("APPDATA", str(Path.home()))) / "Run8DLC"
        alt.mkdir(parents=True, exist_ok=True)
        return alt
    d = base / "data"
    d.mkdir(exist_ok=True)
    for name in ("config.json", "catalog.json", "ledger.json",
                 "mapping.json", "quarantine.json", "transactions.txt",
                 "ocr_cache.json", "ocr_receipts.ps1", "img_convert.ps1",
                 "run8dlc.ico", "report.html", "last_receipt_page.html",
                 "media", "uninstalled"):
        legacy = base / name
        tgt = d / name
        if legacy.exists() and not tgt.exists():
            try:
                legacy.rename(tgt)
            except OSError:
                pass
    return d


APP_DIR = _resolve_app_dir()
DATA_DIR = _resolve_data_dir()
VERSION = "0.9.17"
DEMO_MODE = False

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36 run8dlc/" + VERSION)

def detect_defaults():
    run8 = ""
    for c in (r"C:\Run8Studios\Run8 Train Simulator V3",
              r"D:\Run8Studios\Run8 Train Simulator V3",
              r"C:\Run8Studios\Run8 Train Simulator V2"):
        if Path(c).is_dir():
            run8 = c
            break
    return {
        "installers_dir": str(APP_DIR / "Installers"),
        "transactions_dir": str(DATA_DIR / "Receipts"),
        "backup_dir": str(APP_DIR / "Backups"),
        "run8_install": run8 or r"C:\Run8Studios\Run8 Train Simulator V3",
        "updater_exe": str(Path(run8 or r"C:\Run8Studios\Run8 Train Simulator V3")
                           / "Run8_Updater.exe"),
        "catalog_pages": [
            "https://www.run8studios.com/routes.shtml",
            "https://www.run8studios.com/trainsets.shtml"
        ],
        "store_base": "https://www.3dts-onlinestore.com/"
    }


DEFAULT_CONFIG = detect_defaults()

# ---------------------------------------------------------- embedded assets

CATALOG_SEED = r'''{"catalog_date":"2026-07-10","source_pages":["https://www.run8studios.com/routes.shtml","https://www.run8studios.com/trainsets.shtml"],"products":[{"id":"base_v3","name":"Run8 Train Simulator V3 (base sim)","category":"Base","price":50.0,"url":"https://www.3dts-onlinestore.com/store_run8_v3.php","hints":["run8v3install","run8v3","v3install","install"],"desc":"The Run8 V3 base sim: the Southern California region with the Mojave Sub (including Barstow yard) and the Needles Sub, plus a deep default equipment roster.","info_url":"https://www.run8studios.com/run83.shtml"},{"id":"base_v2","name":"Run8 Train Simulator V2 (legacy base)","category":"Base","price":null,"url":null,"hints":["run8v2"],"desc":"The legacy Run8 V2 base sim. Route and trainset addons are compatible with V2 or V3 unless noted."},{"id":"updates","name":"Run8 Updates / Updaters (free)","category":"Base","price":null,"url":"https://www.run8studios.com/updates_dlc.shtml","hints":["updater","update","southeastregionupdate","alinewaycrossupdate"],"desc":"Free official updates and combined updaters for the sim and its routes."},{"id":"default_equip","name":"Default Equipment Placeholder (free)","category":"Base","price":null,"url":null,"hints":["defaultequipment"],"desc":"The free default-equipment placeholder package."},{"id":"fresno_south","name":"UP Fresno Sub South","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_fresno_south.php","hints":["fresnosouth","fresnosubsouth","upfresno"],"desc":"85 miles of the UP Fresno Sub between CP Saco (Bakersfield) and CP Goble (Fresno), packed with grain facilities and industries. Grade charts and ops notes in the UserGuide.","info_url":"https://www.run8studios.com/fresno_south.shtml"},{"id":"csx_savannah","name":"CSX Savannah Sub","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_csx_savannah.php","hints":["savannah"],"desc":"120 miles linking the Nahunta Sub (A-Line/Waycross) at Ludowici, GA to Yemassee, SC. Savannah and Southover Yards, the Riceboro Southern, 20 industries, and 6 daily passenger trains including the Auto Train.","info_url":"https://www.run8studios.com/savannah.shtml"},{"id":"ags_south1","name":"NS AGS South Phase 1","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_ags_south1.php","hints":["nsags","agssouth","agsphase"],"desc":"Phase 1 of the NS AGS South District: ~85 miles from Norris Yard (Birmingham) to Powers, AL. Hump classification at Norris, BRI intermodal, 11 local jobs, 2 Amtrak stops.","info_url":"https://www.run8studios.com/ags_south1.shtml"},{"id":"pittsburgh","name":"NS Pittsburgh Sub East","category":"Route","price":40.0,"url":"https://www.3dts-onlinestore.com/store_run8_pittsburgh.php","hints":["pittsburgh"],"desc":"60+ miles of NS mainline over the Alleghenies from Johnstown to Tyrone, adding a third track for the climb out of Altoona. Rose and Woodvale Yards, 45 industries, Amtrak.","info_url":"https://www.run8studios.com/pitts.shtml"},{"id":"roseville","name":"SP-UP Roseville Sub","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_roseville.php","hints":["roseville"],"desc":"200+ miles of SP-UP mountain railroading from JR Davis Yard over Donner to Sparks, NV. AI humping at JR Davis, Colfax and Truckee Yards, 70 industries, the Rocklin Rocket.","info_url":"https://www.run8studios.com/roseville.shtml"},{"id":"southfork","name":"NS South Fork Secondary","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_southfork.php","hints":["southfork","nssfs"],"desc":"The 30-mile coal-only NS South Fork Secondary: a 2.4% ruling grade out of South Fork to the Rose Bud, Huskin Run and Shade Creek mines. Push-pull 3x3/3x2 mine runs.","info_url":"https://www.run8studios.com/southfork.shtml"},{"id":"bakersfield","name":"BNSF Bakersfield Sub","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_bakersfield.php","hints":["bakersfield"],"desc":"110 miles of BNSF mainline between Bakersfield Yard and Calwa Yard, Fresno. Grain facilities up the valley and 12 daily Amtrak San Joaquins across 4 stations.","info_url":"https://www.run8studios.com/bakersfield.shtml"},{"id":"baldwin","name":"CSX Baldwin","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_baldwin.php","hints":["baldwin"],"desc":"Fills the missing gap in A-Line operations: the Baldwin Diamond where the Callahan, Jacksonville Terminal, Tallahassee and Wildwood Subs all meet, plus Baldwin Yard.","info_url":"https://www.run8studios.com/baldwin.shtml"},{"id":"modesto","name":"BNSF-UP Fresno to Modesto","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_modesto.php","hints":["fresnomodesto","modesto"],"desc":"250+ miles: BNSF Fresno to Riverbank and UP to Covell, with the M&ET at Modesto and its new MP15. Several yards and four Amtrak stops.","info_url":"https://www.run8studios.com/modesto.shtml"},{"id":"lonepine","name":"Lone Pine Branch & Trona Railway","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_lonepine.php","hints":["lonepine","trona"],"desc":"The UP/SP Lone Pine Branch from Mojave Yard to Searles, plus the 30-mile Trona Railway serving three mineral operations at Searles Lake.","info_url":"https://www.run8studios.com/lonepine.shtml"},{"id":"csx_fitz","name":"CSX Fitzgerald Sub (HyRail)","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_hrs_csx_fitz.php","hints":["fitzgerald","fitz"],"desc":"220 miles of CSX's Fitzgerald Sub from Manchester, GA south to Waycross, set in the busy early 2000s: Florida coal, freight, intermodal, molten sulfur, phosphate and autoracks.","info_url":"https://www.run8studios.com/fitz.shtml"},{"id":"sanberdoo","name":"BNSF San Bernardino Sub","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sanberdoo.php","hints":["sanberdoo","sanbernardino","sanberdo","bnsfsbd","sbdsub"],"desc":"BNSF between San Bernardino/Colton and LA (Hobart Yard), plus the UP Alhambra Sub, the LAJ, and LAUPT for Amtrak.","info_url":"https://www.run8studios.com/sanberdoo.shtml"},{"id":"arvinoak","name":"Arvin & Oak Creek Branches","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_arvinoak.php","hints":["arvinoak","arvin","oakcreek"],"desc":"Two Bakersfield-area branches: the 17-mile Arvin Branch through Lamont to Arvin, and the unsignalled 2%-grade Oak Creek Branch from Mojave to the CalPortland cement plant.","info_url":"https://www.run8studios.com/arvin_oakcreek.shtml"},{"id":"seligman_west","name":"BNSF Seligman Sub West","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_seligman-west.php","hints":["seligman"],"desc":"150 miles of BNSF transcon from Needles, CA to Seligman, AZ along Route 66. Kingman Industrial Park locals and the 90-mph Southwest Chief with its Kingman crew change.","info_url":"https://www.run8studios.com/seligman-west.shtml"},{"id":"waycross","name":"CSX Waycross (HyRail)","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_hrs_csx_waycross.php","hints":["waycross","csxwax"],"desc":"80 miles of southeast Georgia CSX from the A-Line limits to Waycross and Rice Yard's 64-track operational hump bowl, plus the Rayonier pulp plant at Doctortown.","info_url":"https://www.run8studios.com/waycross.shtml"},{"id":"cajon","name":"BNSF Cajon Sub","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8-cajon.php","hints":["cajon"],"desc":"120 miles of BNSF/UP/Amtrak: Barstow to San Bernardino, Silverwood to West Colton, and the 3.4% westbound drop down Cajon Pass. San Bernardino A/B, West Colton and Colton Yards.","info_url":"https://www.run8studios.com/cajon.shtml"},{"id":"csx_aline","name":"CSX A Line (HyRail)","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_hrs_csx-aline.php","hints":["csxaline","aline"],"desc":"HyRail's 150+ miles of CSX A-Line from north of Folkston, GA to south of Orlando: intermodal, autorack loading, unit coal, aggregates and orange juice. Set pre-SunRail.","info_url":"https://www.run8studios.com/aline.shtml"},{"id":"selkirk","name":"CSX Selkirk Terminal","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_selkirk.php","hints":["selkirk"],"desc":"CSX Selkirk, one of the Northeast's largest classification yards, with a working 3.2% hump and bowl. 70+ trains a day originate, terminate or pass through.","info_url":"https://www.run8studios.com/selkirk.shtml"},{"id":"mohawk","name":"CSX Mohawk Sub","category":"Route","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mohawk.php","hints":["mohawk"],"desc":"120 miles of the CSX Mohawk Sub, Syracuse to Amsterdam (merges with Selkirk). Dewitt Yard intermodal, Amtrak with 4 stations, the MHWA and NYSW, accurate signals and defect detectors.","info_url":"https://www.run8studios.com/mohawk.shtml"},{"id":"mp15_pack1","name":"EMD MP15 Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mp15_pack1.php","hints":["mp15pack1","mp1501","mp151"],"family":["mp15"],"desc":"Run8 RR, BNSF Heritage I, Amtrak, UP and CSX YN3 MP15s with 3D cabs, animated windows, doors and wipers."},{"id":"mp15_pack2","name":"EMD MP15 Pack 2","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mp15_pack2.php","hints":["mp15pack2","mp1502","mp152"],"family":["mp15"],"desc":"BN, BNSF Swoosh, NS and SP MP15s with 3D cabs and full animations."},{"id":"mp15_pack3","name":"EMD MP15 Pack 3","category":"Locomotive","price":20.0,"url":"https://www.3dts-onlinestore.com/store_run8_mp15_pack3.php","hints":["mp15pack3","mp1503","mp153"],"family":["mp15"],"desc":"BN-patched BNSF, SP-patched UP, GMTX and CSX MP15s with 3D cabs."},{"id":"dash9_pack1","name":"GE Dash 9 Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_dash9_pack1.php","hints":["dash9pack1","dash901","dash91"],"family":["dash9","c449","c44"],"desc":"Southern Pacific and Santa Fe Dash 9s in fresh and weathered paint, with two cab styles."},{"id":"dash9_pack2","name":"GE Dash 9 Pack 2","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_dash9_pack2.php","hints":["dash9pack2","dash902","dash92"],"family":["dash9","c449","c44"],"desc":"BNSF Warbonnet, BNSF Heritage II, and fresh/weathered UP Dash 9s."},{"id":"dash9_pack3","name":"GE Dash 9 Pack 3","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_dash9_pack3.php","hints":["dash9pack3","dash903","dash93"],"family":["dash9","c449","c44"],"desc":"BNSF Heritage I and Swoosh units, plus CSX YN2 and YN3 Dash 9s."},{"id":"sd70ace_heritage_01","name":"EMD SD70ACe Heritage Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd70ace_heritage_01.php","hints":["sd70aceheritage","heritage01","heritage"],"family":["sd70ace"],"desc":"Six heritage SD70ACes: SP, Rio Grande, WP, Katy, Jersey Central, and the KCS Southern Belle."},{"id":"sd45_2_01","name":"EMD SD45-2 Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd45-2_01.php","hints":["sd45201","sd452pack1","sd4521","sd452pack01"],"family":["sd452","sd45"],"desc":"Eight liveries: ATSF (+Bicentennial), Arizona & California, BNSF, CSX (+Hockey), Conrail and Seaboard System."},{"id":"sd45_2_02","name":"EMD SD45-2 Pack 2","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd45-2_02.php","hints":["sd45202","sd452pack2","sd4522","sd452pack02"],"family":["sd452","sd45"],"desc":"ATSF Kodachrome and Bookend/Pinstripe, Trona Railway, BNSF patched, NS, CSX gray-blue, Conrail Quality and SCL."},{"id":"sd40t2_01","name":"EMD SD40T-2 Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd40t2_01.php","hints":["sd40t2"],"family":["sd40t2","sd40t"],"desc":"Tunnel motors: SP (+patched UP), Rio Grande (+patched UP), and UP SD40T-2s."},{"id":"sd40201","name":"EMD SD40-2 Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd40201.php","hints":["sd40201","sd402pack1","sd4021"],"family":["sd402"],"desc":"BNSF, ATSF, UP, CSX and Run8 Western SD40-2s with four new cabs."},{"id":"sd40202","name":"EMD SD40-2 Pack 2","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd40202.php","hints":["sd40202","sd402pack2","sd4022"],"family":["sd402"],"desc":"SP, ATSF Kodachrome, NS, GCFX and Run8 Western SD40-2s."},{"id":"sd40203","name":"EMD SD40-2 Pack 3","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd40203.php","hints":["sd40203","sd402pack3","sd4023"],"family":["sd402"],"desc":"BNSF Swoosh, BN, Conrail and Cargill AgHorizons SD40-2s."},{"id":"sd40204","name":"EMD SD40-2 Pack 4","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd40204.php","hints":["sd40204","sd402pack4","sd4024"],"family":["sd402"],"desc":"BN patched, ATSF patched, old CSX, and Florida East Coast SD40-2s."},{"id":"sd40205","name":"EMD SD40-2 Pack 5","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_sd40205.php","hints":["sd40205","sd402pack5","sd4025"],"family":["sd402"],"desc":"SOO Line, Chessie System, Seaboard System and Canadian National SD40-2s."},{"id":"gp40201","name":"EMD GP40-2 Pack 1","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_gp40201.php","hints":["gp40201","gp402pack1","gp4021"],"family":["gp402","gp40"],"desc":"Western Pacific, Rio Grande, SP and UP GP40-2s with new cabs."},{"id":"gp40202","name":"EMD GP40-2 Pack 2","category":"Locomotive","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_gp40202.php","hints":["gp40202","gp402pack2","gp4022"],"family":["gp402","gp40"],"desc":"CSX (+hockey stick), Run8 Western, Conrail and BNSF Swoosh GP40-2s."},{"id":"amtrak01","name":"Amtrak Trainsets & GE P42DC Pack 1","category":"Passenger","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_amtrak01.php","hints":["amtrak01","amtrak1","superliner","p42"],"family":["p42"],"desc":"P42s plus Heritage baggage, Superliner transition sleepers, sleepers, coach-baggage, coaches, diners and lounges in Phase III/IV/V, with MHC cars."},{"id":"amfleet2_01","name":"Amtrak Amfleet II Pack 1","category":"Passenger","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_amfleet2_01.php","hints":["amfleet"],"family":["amfleet"],"desc":"Amfleet II coach and lounge cars in Phase III, IV and IVb."},{"id":"viewliner2_pack1","name":"Amtrak Viewliner II Pack 1","category":"Passenger","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_viewliner2_pack1.php","hints":["viewliner"],"family":["viewliner"],"desc":"Viewliner II baggage (P3/P3A), baggage-dorm, sleepers and diner, with 3D diner and sleeper interiors."},{"id":"openhopper01","name":"Open Hopper Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_openhopper01.php","hints":["openhopper01","openhopper1","openhoppers01","openhoppers1","coalballast"],"family":["openhopper"],"desc":"3-bay and 4-bay open hoppers plus three styles of ballast cars, all loadable and unloadable."},{"id":"openhopper02","name":"Open Hopper Pack 2 (Bethgon & Wood Chip)","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_openhopper02.php","hints":["openhopper02","openhopper2","bethgon","woodchip","openhopperspack02","openhopperspack2","openhoppers02"],"family":["bethgon","woodchip","openhopper"],"desc":"Bethgon tubs in BNSF, UP (CMO) and SEMX, plus wood chip cars in CSXT Family Lines, GPSX and IFRX."},{"id":"openhopper03","name":"Open Hopper Pack 3 (AutoFlood III Coal)","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_openhopper03.php","hints":["openhopper03","openhopper3","autoflood","openhopperspack03","openhopperspack3","openhoppers03"],"family":["autoflood","openhopper"],"desc":"Eight aluminum AutoFlood III coal hoppers: BNSF, NS, OUCX, UP, CEFX, GGPX, TILX, MBKX."},{"id":"plastic_pellet01","name":"Plastic Pellet Cars Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_plastic_pellet01.php","hints":["plasticpellet","pellet","pphoppers","pphopper"],"family":["pellet","pphopper"],"desc":"Six ACF5250 plastic pellet hoppers: two ACFX liveries, AMCX, BPRX, ETCX and KCIX."},{"id":"spc509","name":"SP C-50-9 Cabooses","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_spc509.php","hints":["spc509","c509","spc50"],"family":["c509"],"desc":"SP C-50-9 cabooses: Railroad Police, Kodachrome and classic SP, with interiors and a working train brake gauge."},{"id":"atsfce11","name":"ATSF CE11 Cabooses","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_atsfce11.php","hints":["atsfce11","ce11"],"family":["ce11"],"desc":"ATSF CE11 cabooses in Kodachrome, older brown, and the classic scheme, with interiors."},{"id":"mow_pack1","name":"MOW Pack 1 (Tie & Wheel Cars)","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mow_pack1.php","hints":["mowpack1","mow01","mow1","mow"],"family":["tiecar","wheelcar","mowtie"],"desc":"Loaded and empty UP, CSX and BNSF tie cars and wheel cars."},{"id":"wellcar01","name":"Wellcar Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_wellcar01.php","hints":["wellcar01","wellcar1","wellcars01","wellcars1"],"family":["wellcar"],"desc":"Eight smooth-side wellcars: DTTX with Swift and UPS single/double stacks, FEC with trailer loads."},{"id":"wellcar02","name":"Wellcar Pack 2","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_wellcar02.php","hints":["wellcar02","wellcar2","wellcars53"],"family":["wellcar"],"desc":"Twelve wellcars, DTTX and FEC: Hub Group, UMAX, CSX, FedEx and EMP stack configurations."},{"id":"wellcar03","name":"Wellcar Pack 3 (40ft International)","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_wellcar03.php","hints":["wellcar03","wellcar3","wellcars40"],"family":["wellcar"],"desc":"Four 40ft wellcars with twenty-four international container load configurations, plus empties."},{"id":"autorack01","name":"Autorack Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_autorack01.php","hints":["autorack01","autorack1","autoracks01","autoracks1"],"family":["autorack"],"desc":"Four bi-levels (Conrail, CSX, UP, CN) and four tri-levels (BN, NS, SP, BNSF)."},{"id":"autorack02","name":"Autorack Pack 2","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_autorack02.php","hints":["autorack02","autorack2","autoracks02","autoracks2"],"family":["autorack"],"desc":"Six ETTX tri-levels and two TTGX bi-levels: CNW, ATSF, DRGW, CPR, N&W, GT, WP and the Frisco."},{"id":"pig01","name":"Piggyback Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_pig01.php","hints":["piggyback1","pig01","pig1","piggyback01","piggybacks"],"family":["r8pig","kttx","rttx","ttwx","ttax"],"desc":"Nine piggyback flats with UPS, Yellow Freight, Roadway, ABF, Marten, Stevens Transport and Alliance Shippers trailers."},{"id":"pig02","name":"Piggyback Pack 2","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_pig02.php","hints":["piggyback2","pig02","pig2","piggyback02"],"family":["r8pig","kttx","rttx","ttwx","ttax"],"desc":"Ten piggybacks: SP Golden Pig Service, Transamerica, Conrail, BN, Santa Fe, Preferred 45, XTRA, UPSZ/UPOZ, N&AC and Seaboard trailers."},{"id":"reefer01","name":"Refrigerator Car Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_reefer01.php","hints":["reefer","refrigerator"],"family":["reefer"],"desc":"72ft Greenbrier cryogenic, 64ft Millennium TRPX Tropicana, PCF 57 ARMN rebuild, and TrinCool UP/BNSF cars, with sounds when loaded."},{"id":"2baycovhop01","name":"2 Bay Covered Hopper Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_2baycovhop01.php","hints":["2baycovhop01","2baycovhop1","2baycoveredhopper1","2baycoveredhopperpack1","2baycoveredhoppers01","2baycoveredhoppers1"],"family":["acf2970","acf2700"],"desc":"ACF 2970s (BNSF, BN, SP), Trinity 3281s (CEMX, CEFX, TILX, NS), and a CSX ACF 2700."},{"id":"2baycovhop02","name":"2 Bay Covered Hopper Pack 2","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_2baycovhop02.php","hints":["2baycovhop02","2baycovhop2","2baycoveredhopper2","2baycoveredhopperpack2","2baycoveredhoppers02","2baycoveredhoppers2"],"family":["acf2970","acf2700"],"desc":"ACF 2700s (Chessie/CSX, D&TS/CN), ACF 2970s (CNW, N&W, W&W), and Trinity 3281s (GACX, MBKX, MCEX)."},{"id":"gondola01","name":"Gondola Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_gondola01.php","hints":["gondola"],"family":["gondola"],"desc":"Six gon types, empty and loaded: steel coils, ties, rebar, pipes, scrap and wire across EJ&E, UP, CR, MP, SP, GONX, CSX, DJJX, BNSF and IHB."},{"id":"centerbeam01","name":"Centerbeam Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_centerbeam01.php","hints":["centerbeam"],"family":["centerbeam"],"desc":"BC Rail, CN, IC, NOK and two TTZX centerbeams, in empty and loaded versions."},{"id":"fmcboxcars01","name":"FMC Box Cars Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_fmcboxcars01.php","hints":["fmcbox"],"family":["fmc"],"desc":"FMC 5347 boxcars: SP, RailBox, C&G, NOPB, EC&H, MT&W, Cotton Belt and Route Rock."},{"id":"mixed_freight01","name":"Mixed Freight Pack 1","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight01.php","hints":["mixedfreight01","mixedfreight1","mixedfreightpack1","mixedfreightpack01"],"family":["ortner","evanscoil","coilcarevans","gundbox","ps4750"],"desc":"Thrall bulkhead flats with pole, beam and pipe loads, two Greenville 86ft hi-cube auto-parts cars, an FGE reefer, and a BN shoving platform."},{"id":"mixed_freight02","name":"Mixed Freight Pack 2","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight02.php","hints":["mixedfreight02","mixedfreight2","mixedfreightpack2"],"family":["ortner","evanscoil","coilcarevans","gundbox","ps4750"],"desc":"Evans coil cars in BNSF, Conrail and CSS (covered, open and empty), plus five PS4750 3-bay covered hoppers."},{"id":"mixed_freight03","name":"Mixed Freight Pack 3","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight03.php","hints":["mixedfreight03","mixedfreight3","mixedfreightpack3"],"family":["ortner","evanscoil","coilcarevans","gundbox","ps4750"],"desc":"60ft hi-cube boxcars (NS, CSX, TTX), T104 tank cars (ADMX, TILX sulfur), and T389 ACF LPG/NH3 tanks."},{"id":"mixed_freight04","name":"Mixed Freight Pack 4","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight04.php","hints":["mixedfreight04","mixedfreight4","mixedfreightpack4"],"family":["ortner","evanscoil","coilcarevans","gundbox","ps4750"],"desc":"40ft Class K344 Ortner aggregate hoppers (CSX, FEC, SP, Conrail), two Raceland rebuilds, and a white ACFX T389 tank."},{"id":"mixed_freight05","name":"Mixed Freight Pack 5","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight05.php","hints":["mixedfreight05","mixedfreight5","mixedfreightpack5"],"family":["ortner","evanscoil","coilcarevans","gundbox","ps4750"],"desc":"60ft boxcars in WP, CP and UP, plus T104 Trinity tanks in JRSX, TILX sulfur and UTLX."},{"id":"mixed_freight06","name":"Mixed Freight Pack 6","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight06.php","hints":["mixedfreight06","mixedfreight6","mixedfreightpack6"],"family":["mowgeo","hoppercarbon"],"desc":"CSX MoW Geo car, a 60ft HC UP boxcar, CABX/ECQX carbon hoppers, and GATX/UTLX T389 LPG tanks."},{"id":"mixed_freight07","name":"Mixed Freight Pack 7","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight07.php","hints":["mixedfreight07","mixedfreight7","mixedfreightpack7"],"family":["t765","cottonseed"],"desc":"T765 argon tank cars (IAPX, LTCX), GB7 open hoppers (BNSF, UP), and 60ft GC boxcars (FXE, GATX)."},{"id":"mixed_freight08","name":"Mixed Freight Pack 8","category":"Rolling Stock","price":null,"url":"https://www.3dts-onlinestore.com/store_run8_mixed_freight08.php","hints":["mixedfreight08","mixedfreight8","mixedfreightpack8"],"family":["trin4000"],"desc":"PS4750 covered hoppers (BN, ICG, SCL, CSX, PLCX) and a DGHX Trinity 4000 open hopper."},{"id":"hrs_free","name":"HRS free add-ons (free)","category":"Base","price":null,"url":null,"hints":["hrsfree","freeamtrak","freeamtrakautoracks"],"desc":"Free HyRail Simulations add-ons, like the Amtrak autoracks."}]}'''

OCR_PS1 = r'''# run8dlc receipt OCR -- uses the OCR engine built into Windows 10/11.
# Called by: python run8dlc.py ocr-receipts
# Must run under Windows PowerShell 5.x (powershell.exe), not pwsh 7.
param(
    [Parameter(Mandatory=$true)][string]$InDir,
    [Parameter(Mandatory=$true)][string]$OutJson,
    [string[]]$Files
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Runtime.WindowsRuntime

$null = [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.RandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime]

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
    Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
                   $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]

function Await($WinRtTask, $ResultType) {
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}

$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($null -eq $engine) {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage(
        [Windows.Globalization.Language]::new("en-US"))
}
if ($null -eq $engine) {
    Write-Error "No OCR language pack available. Install the English language pack in Windows Settings."
    exit 2
}

if ($Files -and $Files.Count -gt 0) {
    $targets = $Files | ForEach-Object { Get-Item (Join-Path $InDir $_) }
} else {
    $targets = Get-ChildItem -Path $InDir -File |
        Where-Object { $_.Extension -match '\.(png|jpg|jpeg|bmp)$' }
}

$results = [ordered]@{}
$i = 0
foreach ($f in $targets) {
    $i++
    Write-Output ("OCR {0}/{1}  {2}" -f $i, $targets.Count, $f.Name)
    try {
        $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($f.FullName)) ([Windows.Storage.StorageFile])
        $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $bitmap = Await ($decoder.GetSoftwareBitmapAsync(
            [Windows.Graphics.Imaging.BitmapPixelFormat]::Bgra8,
            [Windows.Graphics.Imaging.BitmapAlphaMode]::Premultiplied)) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $ocr = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        $results[$f.Name] = @{ ok = $true; mtime = [long]($f.LastWriteTimeUtc - (Get-Date "1970-01-01")).TotalSeconds; text = $ocr.Text }
        $stream.Dispose()
    } catch {
        $results[$f.Name] = @{ ok = $false; mtime = 0; text = ""; error = $_.Exception.Message }
        Write-Output ("  failed: {0}" -f $_.Exception.Message)
    }
}

$results | ConvertTo-Json -Depth 4 | Out-File -FilePath $OutJson -Encoding UTF8
Write-Output ("wrote {0}" -f $OutJson)
'''

IMG_PS1 = r'''# run8dlc image converter -- JPEG/PNG -> sized PNGs via System.Drawing
param([Parameter(Mandatory=$true)][string]$SrcDir,
      [Parameter(Mandatory=$true)][string]$OutDir,
      [string[]]$NoCrop = @())
$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Drawing
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$files = Get-ChildItem -Path $SrcDir -File | Where-Object { $_.Extension -match "\.(jpe?g|png|gif|bmp)$" }
$i = 0
foreach ($f in $files) {
    $i++
    Write-Output ("IMG {0}/{1}  {2}" -f $i, $files.Count, $f.Name)
    try {
        $img = [System.Drawing.Image]::FromFile($f.FullName)
        $cropY = 0
        $srcH = $img.Height
        if ($NoCrop -notcontains $f.BaseName) {
            $srcH = [int]([double]$img.Height * 0.82)
            $target = [int]([double]$img.Width * 0.52)
            if ($srcH -gt $target) {
                $cropY = [int](($srcH - $target) / 2)
                $srcH = $target
            }
        }
        foreach ($spec in @(@(560, ""), @(168, "_t"))) {
            $w = [int]$spec[0]; $suffix = $spec[1]
            if ($img.Width -lt $w) { $w = $img.Width }
            $h = [int]([double]$srcH * $w / $img.Width)
            $bmp = New-Object System.Drawing.Bitmap($w, $h)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.InterpolationMode = "HighQualityBicubic"
            $dst = New-Object System.Drawing.Rectangle(0, 0, $w, $h)
            $srcR = New-Object System.Drawing.Rectangle(0, $cropY,
                                                        $img.Width, $srcH)
            $g.DrawImage($img, $dst, $srcR,
                         [System.Drawing.GraphicsUnit]::Pixel)
            $g.Dispose()
            $out = Join-Path $OutDir ($f.BaseName + $suffix + ".png")
            $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
            $bmp.Dispose()
        }
        $img.Dispose()
    } catch { Write-Output ("  failed: {0}" -f $_.Exception.Message) }
}
Write-Output "conversion done"
'''


ICON_B64 = "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"


def ensure_assets():
    """First run in a fresh folder: materialize the bundled catalog,
    mapping template, OCR helper, and icon."""
    cat = DATA_DIR / "catalog.json"
    if not cat.exists():
        cat.write_text(CATALOG_SEED, encoding="utf-8")
    mp = DATA_DIR / "mapping.json"
    if not mp.exists():
        save_json(mp, {"__help__": "Manual overrides: 'substring of "
                        "filename' -> product id from catalog.json."})
    ps = DATA_DIR / "ocr_receipts.ps1"
    if not ps.exists() or ps.read_text(encoding="utf-8") != OCR_PS1:
        ps.write_text(OCR_PS1, encoding="utf-8")
    ic = DATA_DIR / "run8dlc.ico"
    if not ic.exists():
        import base64 as _b64
        ic.write_bytes(_b64.b64decode(ICON_B64))
    return cat


EXE_EXTS = {".exe"}
RECEIPT_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".html", ".htm", ".pdf"}

# ---------------------------------------------------------------- utilities

def out(msg=""):
    """ASCII-safe print (old Windows consoles choke on unicode)."""
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", "replace").decode("ascii"))
    except BrokenPipeError:
        pass


def load_json(path, default):
    """Damaged files never take the app down: a copy of the bad file
    is kept beside it and the default is used, so Settings -> Restore
    from backup / Reset to defaults stay reachable."""
    p = Path(path)
    if not p.exists():
        return default
    try:
        with open(p, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        try:
            bad = p.with_suffix(p.suffix + ".corrupt")
            shutil.copy2(p, bad)
            out(f"WARNING: {p.name} is damaged -- continuing with "
                f"defaults (damaged copy kept as {bad.name})")
        except Exception:
            pass
        return default


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        f.write("\n")


def squash(s):
    """lowercase, strip every non-alphanumeric -> canonical comparable form"""
    return re.sub(r"[^a-z0-9]+", "", str(s).lower())


PREFIX_RE = re.compile(
    r"^(run8v3|run8v2|run8|r8v3|r8v2|r82|r8)([_\-. ]+|(?=[a-z0-9]))", re.I)
SUFFIX_RE = re.compile(r"([_\-. ]*(setup|installer|install))+$", re.I)


def core_stem(filename):
    """Strip r8v3_/run8_ prefixes and _setup/_install suffixes from a stem."""
    stem = Path(filename).stem
    m = PREFIX_RE.match(stem)
    # only strip the prefix when a meaningful remainder is left --
    # 'Run8v2' must stay 'Run8v2', not become 'v2'
    s = stem[m.end():] if m and len(stem) - m.end() >= 3 else stem
    s2 = SUFFIX_RE.sub("", s)
    if not squash(s2):          # e.g. 'run8v3_install' -> keep 'install'
        s2 = s
    if not squash(s2):
        s2 = stem
    return s2


def numseq(s):
    return [int(x) for x in re.findall(r"\d+", s)]


def seqscore(a, b):
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return 0.90 + 0.10 * (min(len(a), len(b)) / max(len(a), len(b)))
    return difflib.SequenceMatcher(None, a, b).ratio()


def fetch(url, timeout=30, binary=False):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read()
    if binary:
        return data
    for enc in ("utf-8", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def html_to_text(h):
    h = re.sub(r"(?is)<(script|style).*?</\1>", " ", h)
    h = re.sub(r"(?s)<[^>]+>", " ", h)
    import html as _html
    return re.sub(r"\s+", " ", _html.unescape(h))


# ---------------------------------------------------------------- app state

class App:
    def __init__(self, config_path=None):
        ensure_assets()
        self.config_path = Path(config_path) if config_path else DATA_DIR / "config.json"
        self.config = dict(DEFAULT_CONFIG)
        self.config.update(load_json(self.config_path, {}))
        if not self.config_path.exists():
            save_json(self.config_path, self.config)

        self.catalog_path = DATA_DIR / "catalog.json"
        cat = load_json(self.catalog_path, {"catalog_date": "unknown", "products": []})
        self.catalog = cat
        self.catalog_date = cat.get("catalog_date", "unknown")
        self.products = cat.setdefault("products", [])
        # inherit new fields (descriptions, hint fixes) from the bundled
        # seed without touching user data like prices
        try:
            seed_by_id = {p["id"]: p
                          for p in json.loads(CATALOG_SEED)["products"]}
            for p in self.products:
                s = seed_by_id.get(p["id"])
                if s and s.get("desc") and not p.get("desc"):
                    p["desc"] = s["desc"]
                if s and s.get("info_url") and not p.get("info_url"):
                    p["info_url"] = s["info_url"]
                if s:
                    for h in s.get("hints", []):
                        if h not in p.setdefault("hints", []):
                            p["hints"].append(h)
        except Exception:
            pass
        self.by_id = {p["id"]: p for p in self.products}

        self.mapping_path = DATA_DIR / "mapping.json"
        raw_map = load_json(self.mapping_path, {})
        self.mapping = {k: v for k, v in raw_map.items() if not k.startswith("__")}

        self.ledger_path = DATA_DIR / "ledger.json"
        self.ledger = load_json(self.ledger_path, {"version": 1, "items": []})

        # precompute hint index: DLC hints first (rank 0), Base hints last,
        # longest hint wins within a rank
        self.hint_index = []
        for p in self.products:
            base_rank = 1 if p.get("category") == "Base" else 0
            for h in p.get("hints", []):
                hs = squash(h)
                if hs:
                    self.hint_index.append((base_rank, -len(hs), hs, p["id"]))
        self.hint_index.sort()

    # ------------------------------------------------------------- matching

    def match(self, filename):
        """Return (product_id | None, score, method)."""
        name_sq = squash(Path(filename).name)
        # 1. manual mapping overrides (substring match on full filename)
        for key, pid in self.mapping.items():
            if squash(key) and squash(key) in name_sq and pid in self.by_id:
                return pid, 1.0, "manual"
        stem_sq = squash(core_stem(filename))
        if not stem_sq:
            return None, 0.0, "none"
        # 2. hint substrings (Base category first, then longest hint wins).
        # Base hints (run8v3, update...) only apply when the stem is
        # essentially just that token, so a receipt like
        # "Mixed Freight Pack 5 Run8V3.png" doesn't get eaten by the base sim.
        for _rank, _neglen, hs, pid in self.hint_index:
            if hs in stem_sq:
                if self.by_id[pid].get("category") == "Base" and \
                        len(hs) < 12 and len(stem_sq) > len(hs) + 12:
                    continue
                return pid, 0.99, "hint"
        # 3. fuzzy vs product names, with a number veto: if the filename
        # contains numbers that don't appear in the product name (e.g. a
        # pack index of 2 vs 3), it's almost certainly a different product
        best = (0.0, None)
        fnums = set(numseq(stem_sq))
        for p in self.products:
            cand = squash(p["name"])
            sc = seqscore(stem_sq, cand)
            pnums = set(numseq(cand))
            if fnums and pnums and (fnums - pnums):
                sc *= 0.4
            if sc > best[0]:
                best = (sc, p["id"])
        if best[0] >= 0.75:
            return best[1], best[0], "fuzzy"
        return None, best[0], "none"

    def suggest(self, filename, n=3):
        stem_sq = squash(core_stem(filename))
        scored = []
        for p in self.products:
            sc = seqscore(stem_sq, squash(p["name"]))
            scored.append((sc, p["name"]))
        scored.sort(reverse=True)
        return [f"{name} ({sc:.2f})" for sc, name in scored[:n] if sc > 0.3]

    # ------------------------------------------------------------ scanning

    def scan_installers(self):
        d = Path(self.config["installers_dir"])
        if not d.is_dir():
            return []
        return sorted(f for f in d.iterdir()
                      if f.is_file() and f.suffix.lower() in EXE_EXTS)

    def scan_receipts(self):
        d = Path(self.config["transactions_dir"])
        if not d.is_dir():
            return []
        return sorted(f for f in d.iterdir()
                      if f.is_file() and f.suffix.lower() in RECEIPT_EXTS)

    def scan_game(self, max_depth=3):
        """Harvest directory names inside the Run8 install; hint-match only
        (fuzzy on arbitrary folder names is too noisy)."""
        root = Path(self.config["run8_install"])
        found = {}
        if not root.is_dir():
            return found
        root_depth = len(root.parts)
        for dirpath, dirnames, _files in os.walk(root):
            depth = len(Path(dirpath).parts) - root_depth
            if depth >= max_depth:
                dirnames[:] = []
                continue
            for dn in dirnames:
                dsq = squash(dn)
                if not dsq:
                    continue
                for _rank, _neglen, hs, pid in self.hint_index:
                    prod = self.by_id[pid]
                    if prod.get("category") == "Base":
                        continue
                    if len(hs) >= 5 and hs in dsq:
                        found.setdefault(pid, []).append(
                            str(Path(dirpath, dn).relative_to(root)))
                        break
        return found

    def scan_equipment(self, eligible_pids):
        """Filename-based family detection inside Content\\V3RailVehicles\\Body.
        Equipment DLC installs flat files there (no per-product folders), so
        we look for family tokens (e.g. 'sd402', 'wellcar') in filenames.
        Only checks eligible (already-owned) products: base-sim default
        equipment shares that folder, so file presence alone must never
        create ownership."""
        body = Path(self.config["run8_install"]) / "Content" / \
            "V3RailVehicles" / "Body"
        found = {}
        if not body.is_dir():
            return found
        try:
            names = [f.name for f in body.iterdir() if f.is_file()]
        except OSError:
            return found
        squashed = [(squash(n), n) for n in names]
        for p in self.products:
            pid = p["id"]
            if pid not in eligible_pids:
                continue
            for tok in p.get("family", []):
                ts = squash(tok)
                if not ts:
                    continue
                hit = next((orig for sq, orig in squashed if ts in sq), None)
                if hit:
                    found[pid] = f"Body\\{hit} ('{tok}' family)"
                    break
        return found

    # ------------------------------------------------------------ ownership

    def build_state(self, include_game_scan=True):
        """Structured evidence per product:
        {pid: {"ledger": [items], "installers": [Path], "receipts": [Path],
               "game": [relpath strings], "game_files": [descriptions]}},
        plus unmatched files. "game" = whole directories (safe to
        quarantine); "game_files" = family-level file evidence in shared
        folders (never touched by uninstall)."""
        state = {}
        unmatched = []

        def slot(pid):
            return state.setdefault(pid, {"ledger": [], "installers": [],
                                          "receipts": [], "game": [],
                                          "game_files": []})

        for item in self.ledger.get("items", []):
            pid = item.get("product_id")
            if pid in self.by_id:
                slot(pid)["ledger"].append(item)

        for f in self.scan_installers():
            pid, _sc, _m = self.match(f.name)
            if pid:
                slot(pid)["installers"].append(f)
            else:
                unmatched.append(("installer", f.name, self.suggest(f.name)))

        for f in self.scan_receipts():
            pid, _sc, _m = self.match(f.name)
            if pid:
                slot(pid)["receipts"].append(f)
            else:
                unmatched.append(("receipt", f.name, self.suggest(f.name)))

        if include_game_scan:
            for pid, paths in self.scan_game().items():
                slot(pid)["game"].extend(paths)
            owned_pids = {pid for pid, ev in state.items()
                          if ev["ledger"] or ev["installers"] or ev["receipts"]}
            for pid, desc in self.scan_equipment(owned_pids).items():
                slot(pid)["game_files"].append(desc)

        return state, unmatched

    def build_ownership(self, include_game_scan=True):
        """Returns (owned: {pid: [evidence,...]}, unmatched) -- flat-string
        view of build_state, kept for report/migrate."""
        state, unmatched = self.build_state(include_game_scan)
        owned = {}
        for pid, ev in state.items():
            lst = ["ledger"] * len(ev["ledger"])
            lst += [f"installer: {p.name}" for p in ev["installers"]]
            lst += [f"receipt: {p.name}" for p in ev["receipts"]]
            if ev["game"]:
                lst.append(f"in-game: {ev['game'][0]}")
            elif ev["game_files"]:
                lst.append(f"in-game files: {ev['game_files'][0]}")
            if lst:
                owned[pid] = lst
        return owned, unmatched


# --------------------------------------------------- install lifecycle layer

QUAR_PATH = DATA_DIR / "quarantine.json"
UNINSTALLED_DIR = DATA_DIR / "uninstalled"
# never quarantine these top-level-ish folders even if a hint matches them
PROTECTED_DIR_NAMES = {"content", "regions", "avatars", "sounds", "data",
                       "routes", "trains", "equipment", "config"}


def load_quarantine():
    return load_json(QUAR_PATH, {"version": 1, "records": []})


def save_quarantine(q):
    save_json(QUAR_PATH, q)


def quarantined_pids(q=None):
    q = q or load_quarantine()
    return {r["pid"] for r in q.get("records", [])}


def product_status(pid, state, quar_pids):
    """installed / quarantined / owned / missing"""
    ev = state.get(pid)
    if ev and (ev["game"] or ev.get("game_files")):
        return "installed"
    if pid in quar_pids:
        return "quarantined"
    if ev and (ev["ledger"] or ev["installers"] or ev["receipts"]):
        return "owned"
    return "missing"


def find_installer_for(app, pid):
    """Newest installer EXE in Installers/ that matches this product."""
    best = None
    for f in app.scan_installers():
        got, _sc, _m = app.match(f.name)
        if got == pid:
            if best is None or f.stat().st_mtime > best.stat().st_mtime:
                best = f
    return best


def resolve_product(app, query):
    """Turn a user-supplied name/id fragment into a pid.
    Returns (pid | None, candidate_names)."""
    if query in app.by_id:
        return query, []
    qsq = squash(query)
    if qsq:
        exact = [p for p in app.products if squash(p["name"]) == qsq]
        if len(exact) == 1:
            return exact[0]["id"], []
        subs = [p for p in app.products
                if qsq in squash(p["name"]) or qsq in p["id"]]
        if len(subs) == 1:
            return subs[0]["id"], []
        if len(subs) > 1:
            return None, [p["name"] for p in subs[:8]]
    pid, sc, _m = app.match(query)
    if pid and sc >= 0.75:
        return pid, []
    return None, app.suggest(query)


def launch(path, cwd=None):
    """Start an EXE. Returns True if actually launched (Windows),
    False on other platforms (prints what it would do)."""
    path = Path(path)
    if sys.platform != "win32":
        out(f"  (not Windows -- would launch: {path}"
            + (f"  [cwd {cwd}]" if cwd else "") + ")")
        return False
    try:
        subprocess.Popen([str(path)], cwd=str(cwd) if cwd else None)
    except OSError:
        # elevation-required EXEs raise WinError 740; startfile shows UAC
        os.startfile(str(path))  # noqa: attribute exists on Windows
    return True


def updater_path(app):
    p = app.config.get("updater_exe")
    if p:
        return Path(p)
    return Path(app.config["run8_install"]) / "Run8_Updater.exe"


def _game_paths_for(app, pid, state=None):
    """Validated absolute in-game folders safe to quarantine."""
    if state is None:
        state, _ = app.build_state(include_game_scan=True)
    ev = state.get(pid)
    root = Path(app.config["run8_install"]).resolve()
    good, skipped = [], []
    for rel in (ev["game"] if ev else []):
        ap = (root / rel).resolve()
        try:
            relp = ap.relative_to(root)
        except ValueError:
            skipped.append((rel, "outside the Run8 install"))
            continue
        if len(relp.parts) < 2:
            skipped.append((rel, "too close to the install root"))
            continue
        if squash(ap.name) in PROTECTED_DIR_NAMES:
            skipped.append((rel, "protected folder name"))
            continue
        if not ap.is_dir():
            skipped.append((rel, "not found / not a folder"))
            continue
        if ap not in good:
            good.append(ap)
    return good, skipped


def uninstall_product(app, pid, apply=False, state=None):
    """Move a product's detected in-game folders into ./uninstalled/
    (reversible). Returns (ok, lines)."""
    lines = []
    prod = app.by_id.get(pid)
    if not prod:
        return False, [f"unknown product id: {pid}"]
    good, skipped = _game_paths_for(app, pid, state)
    for rel, why in skipped:
        lines.append(f"  skipping {rel}  ({why})")
    if not good:
        lines.append(f"No in-game folders located for {prod['name']}.")
        lines.append("Detection currently keys off folder names; equipment "
                     "packs may not be locatable until snapshot tuning.")
        return False, lines
    if not apply:
        lines.append(f"DRY RUN -- would quarantine for {prod['name']}:")
        for ap in good:
            lines.append(f"  {ap}")
        lines.append("Re-run with --apply to move these into ./uninstalled/")
        return True, lines
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest_root = UNINSTALLED_DIR / pid / stamp
    dest_root.mkdir(parents=True, exist_ok=True)
    rec = {"pid": pid, "name": prod["name"], "when": stamp, "items": []}
    for i, ap in enumerate(good):
        dst = dest_root / f"{i:02d}_{ap.name}"
        shutil.move(str(ap), str(dst))
        rec["items"].append({"orig": str(ap), "moved_to": str(dst)})
        lines.append(f"  moved {ap.name} -> {dst}")
    q = load_quarantine()
    q["records"].append(rec)
    save_quarantine(q)
    lines.append(f"{prod['name']} quarantined ({len(good)} folder(s)). "
                 "Use 'restore' to undo.")
    return True, lines


def restore_product(app, pid):
    """Move quarantined folders back. Returns (ok, lines)."""
    lines = []
    q = load_quarantine()
    recs = [r for r in q.get("records", []) if r["pid"] == pid]
    if not recs:
        return False, [f"nothing quarantined for '{pid}'"]
    ok = True
    for rec in recs:
        for item in rec["items"]:
            src, dst = Path(item["moved_to"]), Path(item["orig"])
            if not src.exists():
                lines.append(f"  missing in quarantine: {src}")
                ok = False
                continue
            if dst.exists():
                lines.append(f"  target already exists, left in place: {dst}")
                ok = False
                continue
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst))
            lines.append(f"  restored {dst}")
    if ok:
        q["records"] = [r for r in q.get("records", []) if r["pid"] != pid]
        save_quarantine(q)
        name = app.by_id.get(pid, {}).get("name", pid)
        lines.append(f"{name} restored.")
    else:
        lines.append("Some items could not be restored; quarantine record "
                     "kept. Fix manually and re-run.")
    return ok, lines


# ------------------------------------------------------------------ report

def price_str(p):
    return f"${p['price']:.2f}" if p.get("price") is not None else "$?"


def size_str(n):
    """MB below ~1 GB, GB with two decimals above."""
    return (f"{n / 1073741824:.2f} GB" if n >= 1000 * 1048576
            else f"{n / 1048576:.0f} MB")


def cmd_report(app, args):
    owned, unmatched = app.build_ownership(include_game_scan=not args.no_game_scan)

    dlc = [p for p in app.products if p.get("category") not in ("Base",)]
    missing = [p for p in dlc if p["id"] not in owned]
    owned_dlc = [p for p in dlc if p["id"] in owned]

    out(f"run8dlc report  (catalog dated {app.catalog_date}, "
        f"{len(dlc)} DLC products known)")
    out("-" * 70)
    out(f"OWNED:   {len(owned_dlc)} DLC products")
    out(f"MISSING: {len(missing)} DLC products")
    if unmatched:
        out(f"REVIEW:  {len(unmatched)} files did not match anything "
            f"(see report.html / --verbose)")
    out("")

    if missing:
        out("You do not appear to own:")
        total = 0.0
        total_known = True
        by_cat = {}
        for p in missing:
            by_cat.setdefault(p["category"], []).append(p)
        for cat in sorted(by_cat):
            out(f"  [{cat}]")
            for p in sorted(by_cat[cat], key=lambda x: x["name"]):
                out(f"    {p['name']:<45} {price_str(p):>8}")
                if p.get("price") is not None:
                    total += p["price"]
                else:
                    total_known = False
        approx = "" if total_known else " (some prices unknown - run: refresh --prices)"
        out(f"\n  Cost to complete the collection: ${total:.2f}{approx}")
    else:
        out("You own everything currently in the store. Nice.")

    if args.verbose and unmatched:
        out("\nUnmatched files (add overrides to mapping.json):")
        for kind, name, sugg in unmatched:
            out(f"  [{kind}] {name}")
            for s in sugg:
                out(f"      maybe: {s}")

    n_priced = sum(1 for p in dlc if p.get("price") is not None)
    if n_priced < len(dlc) / 2:
        out("\nTip: most prices are unknown. Run:  python run8dlc.py refresh --prices")

    html_path = DATA_DIR / "report.html"
    write_html_report(app, owned, missing, owned_dlc, unmatched, html_path)
    out(f"\nFull report: {html_path}")
    if args.open and sys.platform == "win32":
        os.startfile(html_path)  # noqa


def write_html_report(app, owned, missing, owned_dlc, unmatched, path):
    import html as H

    def esc(s):
        return H.escape(str(s))

    total = sum(p["price"] for p in missing if p.get("price") is not None)
    unknown_prices = any(p.get("price") is None for p in missing)

    rows_missing = ""
    for p in sorted(missing, key=lambda x: (x["category"], x["name"])):
        link = (f'<a href="{esc(p["url"])}" target="_blank">buy</a>'
                if p.get("url") else "")
        rows_missing += (f"<tr><td>{esc(p['category'])}</td>"
                         f"<td>{esc(p['name'])}</td>"
                         f"<td class='r'>{esc(price_str(p))}</td>"
                         f"<td>{link}</td></tr>\n")

    rows_owned = ""
    for p in sorted(owned_dlc, key=lambda x: (x["category"], x["name"])):
        raw = owned.get(p["id"], [])
        ev_list = [e for e in raw if e != "ledger"]
        if len(ev_list) != len(raw):
            ev_list.append("ledger")
        ev_list.sort(key=lambda s: 0 if s.startswith("in-game") else 1)
        ev = "; ".join(ev_list)
        rows_owned += (f"<tr><td>{esc(p['category'])}</td>"
                       f"<td>{esc(p['name'])}</td>"
                       f"<td class='ev'>{esc(ev)}</td></tr>\n")

    rows_review = ""
    for kind, name, sugg in unmatched:
        rows_review += (f"<tr><td>{esc(kind)}</td><td>{esc(name)}</td>"
                        f"<td class='ev'>{esc(', '.join(sugg))}</td></tr>\n")

    doc = f"""<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Run8 DLC Report</title><style>
body{{background:#141414;color:#ddd;font:14px/1.5 'Segoe UI',sans-serif;
     max-width:960px;margin:24px auto;padding:0 16px}}
h1{{font-size:20px;color:#fff}} h2{{font-size:16px;color:#e8b34b;
     border-bottom:1px solid #333;padding-bottom:4px;margin-top:32px}}
table{{border-collapse:collapse;width:100%}}
td,th{{padding:5px 10px;border-bottom:1px solid #2a2a2a;text-align:left;
     vertical-align:top}}
th{{color:#999;font-weight:600}} .r{{text-align:right}}
.ev{{color:#888;font-size:12px}}
a{{color:#6fb3ff}} .tot{{color:#e8b34b;font-weight:700}}
.sub{{color:#888}}
</style></head><body>
<h1>Run8 DLC Report</h1>
<p class="sub">Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} -
catalog dated {esc(app.catalog_date)} -
{len(owned_dlc)} owned / {len(missing)} missing of
{len(owned_dlc) + len(missing)} DLC products</p>

<h2>Missing ({len(missing)})</h2>
<table><tr><th>Category</th><th>Product</th><th>Price</th><th></th></tr>
{rows_missing}
<tr><td></td><td class="tot">Total to complete</td>
<td class="r tot">${total:.2f}{'+' if unknown_prices else ''}</td><td></td></tr>
</table>

<h2>Owned ({len(owned_dlc)})</h2>
<table><tr><th>Category</th><th>Product</th><th>Evidence</th></tr>
{rows_owned}</table>

<h2>Needs review ({len(unmatched)})</h2>
<p class="sub">These files didn't match any product. Add a line to
mapping.json ("filename substring": "product_id") and re-run.</p>
<table><tr><th>Type</th><th>File</th><th>Best guesses</th></tr>
{rows_review}</table>
</body></html>"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


# ----------------------------------------------------------------- migrate

def cmd_migrate(app, args):
    installers = app.scan_installers()
    receipts = app.scan_receipts()
    out(f"Scanning {len(installers)} installers and {len(receipts)} receipts...")

    per_product = {}
    unmatched = []

    for f in installers:
        pid, sc, method = app.match(f.name)
        if pid:
            per_product.setdefault(pid, {"exes": [], "receipts": []})["exes"].append(f)
        else:
            unmatched.append(("installer", f))

    for f in receipts:
        pid, sc, method = app.match(f.name)
        if pid:
            per_product.setdefault(pid, {"exes": [], "receipts": []})["receipts"].append(f)
        else:
            unmatched.append(("receipt", f))

    existing = {i.get("product_id") for i in app.ledger.get("items", [])}
    added = 0
    for pid, files in sorted(per_product.items()):
        if pid in existing:
            continue
        prod = app.by_id[pid]
        newest = None
        for f in files["exes"] + files["receipts"]:
            m = datetime.fromtimestamp(f.stat().st_mtime)
            newest = max(newest, m) if newest else m
        app.ledger["items"].append({
            "product_id": pid,
            "name": prod["name"],
            "category": prod["category"],
            "exe": files["exes"][-1].name if files["exes"] else None,
            "receipt": files["receipts"][-1].name if files["receipts"] else None,
            "transaction_id": None,
            "date": newest.date().isoformat() if newest else None,
            "source": "migrated",
        })
        added += 1

    save_json(app.ledger_path, app.ledger)
    write_transactions_txt(app)
    out(f"Ledger: {added} products added "
        f"({len(app.ledger['items'])} total) -> {app.ledger_path}")

    if unmatched:
        out(f"\n{len(unmatched)} files need review:")
        for kind, f in unmatched:
            out(f"  [{kind}] {f.name}")
            for s in app.suggest(f.name):
                out(f"      maybe: {s}")
        out("\nFix by adding entries to mapping.json, then re-run migrate.")

    if args.organize:
        plan = []
        for item in app.ledger["items"]:
            r = item.get("receipt")
            if not r:
                continue
            src = Path(app.config["transactions_dir"]) / r
            want = re.sub(r'[<>:"/\\|?*]', "", item["name"]) + src.suffix.lower()
            if src.name != want and src.exists():
                dst = src.with_name(want)
                n = 2
                while dst.exists() and dst != src:
                    dst = src.with_name(f"{Path(want).stem} #{n}{src.suffix.lower()}")
                    n += 1
                plan.append((src, dst, item))
        if not plan:
            out("\n--organize: receipt names already consistent.")
        elif args.apply:
            for src, dst, item in plan:
                src.rename(dst)
                item["receipt"] = dst.name
                out(f"  renamed: {src.name} -> {dst.name}")
            save_json(app.ledger_path, app.ledger)
            out(f"\n--organize: {len(plan)} receipts renamed.")
        else:
            out(f"\n--organize DRY RUN ({len(plan)} renames, add --apply to execute):")
            for src, dst, _ in plan:
                out(f"  {src.name}  ->  {dst.name}")


# --------------------------------------------------------------------- add

TXID_URL_KEYS = ("transid", "txid", "tx", "txn", "transactionid", "transaction_id",
                 "transaction", "orderid", "order_id", "order", "id", "token")


def txid_from_url(url):
    """Receipt-page URLs carry the transaction id as a query param."""
    try:
        q = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
    except Exception:
        return None
    low = {k.lower(): v for k, v in q.items()}
    for key in TXID_URL_KEYS:
        for v in low.get(key, []):
            v = v.strip()
            if len(v) >= 6 and re.fullmatch(r"[A-Za-z0-9\-_.]+", v):
                return v
    # fall back: any query value that looks like a PayPal-style token
    for vs in low.values():
        for v in vs:
            if re.fullmatch(r"[A-Z0-9]{12,20}", v.strip()):
                return v.strip()
    return None


def txid_from_page(page_text):
    m = re.search(r"Transaction\s*(?:ID)?\s*[:#]?\s*([A-Za-z0-9\-]{6,})",
                  page_text, re.I)
    return m.group(1) if m else None


def find_download_link(page_html, page_url):
    """The 3DTS receipt page prints the link as PLAIN TEXT after a
    'DL LINK:' label -- no anchor tag. Check that first, then score every
    anchor and every bare URL in the text by download-ness."""
    text = html_to_text(page_html)
    m = re.search(r"DL\s*LINK\s*:?\s*(https?://[^\s<>\"']+)", text, re.I)
    if m:
        return m.group(1).rstrip(".,;)]'\"")
    cands = re.findall(
        r"""<a[^>]*?href\s*=\s*["']?([^"'\s>]+)[^>]*>(.*?)</a>""",
        page_html, re.I | re.S)
    # bare URLs in the visible text count too (no anchor text)
    cands += [(u.rstrip(".,;)]'\""), "") for u in
              re.findall(r"(https?://[^\s<>\"']+)", text, re.I)]
    best, best_score = None, 0
    for href, anchor_text in cands:
        if href.lower().startswith(("javascript:", "mailto:", "#")):
            continue
        h = href.lower()
        t = re.sub(r"<[^>]+>", " ", anchor_text).lower()
        score = 0
        if ".exe" in h:
            score += 4
        if "download" in t or "download" in h:
            score += 3
        if re.search(r"file|fetch|transid|(?<![a-z])dl", h):
            score += 2
        if "?" in h:
            score += 1
        if score > best_score:
            best, best_score = href, score
    return urllib.parse.urljoin(page_url, best) if best else None


def filename_from_response(url, headers, fallback_stem=None):
    """Real installer name comes from Content-Disposition when the URL is
    a script like download.php?transid=..."""
    cd = headers.get("Content-Disposition") or ""
    m = re.search(r"""filename\*?\s*=\s*(?:UTF-8'')?["']?([^"';\r\n]+)""",
                  cd, re.I)
    if m:
        name = Path(urllib.parse.unquote(m.group(1).strip())).name
        if name:
            return name
    name = Path(urllib.parse.urlparse(url).path).name
    if name and "." in name and not name.lower().endswith(
            (".php", ".asp", ".cgi")):
        return name
    return (fallback_stem or "download") + ".exe"


def download(url, dest_dir, fallback_stem=None):
    """Download with newline-terminated progress lines. Lines start with
    '[dl]' so the GUI can drive a progress bar off them."""
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=60) as r:
        fname = filename_from_response(url, r.headers, fallback_stem)
        dest = dest_dir / fname
        with open(dest, "wb") as f:
            total = int(r.headers.get("Content-Length") or 0)
            got, next_pct, next_mb = 0, 5, 4 * 1024 * 1024
            while True:
                chunk = r.read(1024 * 256)
                if not chunk:
                    break
                f.write(chunk)
                got += len(chunk)
                if total:
                    pct = got * 100 // total
                    if pct >= next_pct:
                        out(f"[dl] {pct}%  "
                            f"{got/1048576:.1f}/{total/1048576:.1f} MB")
                        next_pct = pct - pct % 5 + 5
                elif got >= next_mb:
                    out(f"[dl] {got/1048576:.1f} MB...")
                    next_mb += 4 * 1024 * 1024
    out(f"[dl] 100%  done: {dest.name} ({dest.stat().st_size/1048576:.1f} MB)")
    return dest


def safe_name(s):
    return re.sub(r'[<>:"/\\|?*]', "", s).strip()


def cmd_add(app, args):
    inst_dir = Path(app.config["installers_dir"])
    inst_dir.mkdir(parents=True, exist_ok=True)

    exe_path, dl_url, page = None, None, None
    txid = args.txid or (txid_from_url(args.url) if args.url else None)
    if txid:
        out(f"  Transaction ID: {txid}")

    if args.url:
        out(f"Fetching receipt page: {args.url}")
        try:
            page = fetch(args.url)
        except Exception as e:
            out(f"  WARNING: could not fetch page ({e}) -- it has probably "
                "expired.")
            if not (args.exe or txid):
                out("  Nothing to record. Re-run with --exe <downloaded "
                    "file> to log this purchase anyway.")
                return 1
        if page:
            if not txid:
                txid = txid_from_page(html_to_text(page))
                if txid:
                    out(f"  Transaction ID: {txid}")
            dl_url = find_download_link(page, args.url)
            if not dl_url and txid:
                dl_url = ("http://www.run8-services.com/download.php"
                          f"?transid={txid}")
                out(f"  no link on the page; constructing the standard one: "
                    f"{dl_url}")
            if dl_url and not args.exe:
                out(f"  Download link: {dl_url}")
                try:
                    exe_path = download(dl_url, inst_dir,
                                        fallback_stem=("purchase_" + txid)
                                        if txid else None)
                except Exception as e:
                    out(f"  ERROR: download failed ({e})")
                    exe_path = None
            elif not dl_url:
                dbg = DATA_DIR / "last_receipt_page.html"
                dbg.write_text(page, encoding="utf-8")
                out("  WARNING: no download link found on the page.")
                out(f"  Saved the raw page to {dbg} -- send it to Claude "
                    "and the parser gets fixed in one pass.")
            if args.dump:
                dbg = DATA_DIR / "last_receipt_page.html"
                dbg.write_text(page, encoding="utf-8")
                out(f"  Raw page saved to {dbg}")

    link = getattr(args, "link", None)
    if not link and txid and not args.url and not args.exe:
        link = ("http://www.run8-services.com/download.php"
                f"?transid={txid}")
        out(f"  constructing the standard download link: {link}")
    if link and exe_path is None and not args.exe:
        out(f"  Download link: {link}")
        dl_url = link
        try:
            exe_path = download(link, inst_dir,
                                fallback_stem=("purchase_" + txid)
                                if txid else None)
        except Exception as e:
            out(f"  ERROR: download failed ({e})")
            exe_path = None

    if args.exe:
        src = Path(args.exe)
        if not src.exists():
            out(f"ERROR: {src} not found")
            return 1
        dest = inst_dir / src.name
        if src.resolve() != dest.resolve():
            shutil.copy2(src, dest)
            out(f"  copied {src.name} -> {inst_dir}")
        exe_path = dest

    # which product? the EXE filename is the strongest signal
    pid = None
    if args.name:
        pid, _sc, _m = app.match(args.name)
    if not pid and exe_path is not None:
        pid, _sc, _m = app.match(exe_path.name)
    prod = app.by_id.get(pid)
    disp = (prod["name"] if prod else
            args.name or (exe_path.stem if exe_path else
                          f"Unknown purchase ({txid or 'no txid'})"))
    out(f"  Product: {disp}" + ("" if prod else "  (no catalog match -- "
        "add a mapping.json override and re-run migrate)"))

    # optional manual receipt attach (legacy PNG workflow)
    receipt_name = None
    if args.receipt:
        src = Path(args.receipt)
        if src.exists():
            tx_dir = Path(app.config["transactions_dir"])
            tx_dir.mkdir(parents=True, exist_ok=True)
            dest = tx_dir / (safe_name(disp) + src.suffix.lower())
            shutil.copy2(src, dest)
            receipt_name = dest.name
            out(f"  Receipt copied -> {dest}")

    items = app.ledger.setdefault("items", [])
    entry = None
    if txid:
        entry = next((i for i in items
                      if (i.get("transaction_id") or i.get("txid")) == txid),
                     None)
    fresh = {
        "product_id": pid,
        "name": disp,
        "category": prod["category"] if prod else None,
        "exe": exe_path.name if exe_path else None,
        "receipt": receipt_name,
        "transaction_id": txid,
        "download_url": dl_url,
        "receipt_page_url": args.url,
        "date": datetime.now().strftime("%Y-%m-%d"),
        "source": "purchase",
    }
    if entry:
        for k, v in fresh.items():
            if v is not None or entry.get(k) is None:
                if k == "date":
                    continue
                entry[k] = v if v is not None else entry.get(k)
        out(f"  Updated the existing ledger entry for txid {txid}")
    else:
        items.append(fresh)
        out(f"  Logged to ledger + transactions.txt")
    save_json(app.ledger_path, app.ledger)
    write_transactions_txt(app)

    if args.install and exe_path:
        out(f"  launching installer: {exe_path.name}")
        launch(exe_path, cwd=exe_path.parent)
    elif args.install:
        out("  (--install skipped: no EXE was obtained)")
    return 0


def write_transactions_txt(app):
    """Human-readable purchase record regenerated from the ledger."""
    items = sorted(app.ledger.get("items", []),
                   key=lambda i: (i.get("date") or "", i.get("name") or ""))
    lines = [f"run8dlc transactions -- regenerated "
             f"{datetime.now().strftime('%Y-%m-%d %H:%M')}",
             f"{len(items)} entries", "",
             f"{'date':<12}{'product':<42}{'transaction id':<22}"
             f"{'installer':<34}source"]
    lines.append("-" * len(lines[-1]))
    for i in items:
        lines.append(f"{(i.get('date') or '?'):<12}"
                     f"{(i.get('name') or '?')[:40]:<42}"
                     f"{(i.get('transaction_id') or i.get('txid') or '-')[:20]:<22}"
                     f"{(i.get('exe') or '-')[:32]:<34}"
                     f"{i.get('source') or '-'}")
    dest = DATA_DIR / "transactions.txt"
    dest.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return dest


def cmd_transactions(app, args):
    dest = write_transactions_txt(app)
    for ln in dest.read_text(encoding="utf-8").splitlines():
        out(ln)
    out(f"\n(saved at {dest})")
    return 0



NAV_NOISE = re.compile(
    r"HOME\s*\||PRESS RELEASES|USA PRODUCTS|UK PRODUCTS|ONLINE STORE|"
    r"CONTACT US|COPYRIGHT|ALL RIGHTS RESERVED|TRAIN SIMULATOR BY RUN8|"
    r"OVERVIEW\s*\||\|\s*UPDATES|meta-Description|^title:|BUY IT NOW|"
    r"CLICK HERE|Learn more|^-{2,}$", re.I)


def product_page_text(url):
    """Fetch a run8studios.com product page and reduce it to readable
    prose for the in-app viewer."""
    h = fetch(url, timeout=15)
    t = html_to_text(h)
    lines, blank = [], 0
    for ln in t.splitlines():
        ln = ln.strip()
        if NAV_NOISE.search(ln):
            continue
        if not ln:
            blank += 1
            if blank > 1:
                continue
        else:
            blank = 0
        lines.append(ln)
    body = "\n".join(lines).strip()
    body = re.sub(r"\n{3,}", "\n\n", body)
    return body or "(page was empty)"



def quarantine_size(app):
    qdir = DATA_DIR / "uninstalled"
    n, total = 0, 0
    if qdir.exists():
        for f in qdir.rglob("*"):
            if f.is_file():
                n += 1
                total += f.stat().st_size
    return n, total


def cmd_purge_quarantine(app, args):
    qdir = DATA_DIR / "uninstalled"
    quar = load_json(DATA_DIR / "quarantine.json", {"items": []})
    n_items = len(quar.get("items", []))
    n_files, total = quarantine_size(app)
    if n_files == 0 and n_items == 0:
        out("Quarantine is already empty.")
        return 0
    out(f"Quarantine holds {n_items} product(s), {n_files} file(s), "
        f"{total/1048576:.0f} MB.")
    if not args.yes:
        out("This PERMANENTLY deletes them -- Restore will no longer be "
            "possible.\nRe-run with --yes to confirm.")
        return 1
    if qdir.exists():
        shutil.rmtree(qdir, ignore_errors=True)
    save_json(DATA_DIR / "quarantine.json", {"items": []})
    out(f"Purged. {total/1048576:.0f} MB freed. Reinstalling any of these "
        "routes now requires running their installers again.")
    return 0



def cmd_backup(app, args):
    """One compressed Run8DLC_Backup.zip -- installers, ledger, records.
    Written to a .part file first, then swapped in, so an interrupted
    backup never destroys the previous good zip."""
    import zipfile
    dest = Path(args.dest or app.config.get("backup_dir")
                or DEFAULT_CONFIG["backup_dir"]).expanduser()
    dest.mkdir(parents=True, exist_ok=True)
    zpath = dest / "Run8DLC_Backup.zip"
    tmp = dest / "Run8DLC_Backup.zip.part"
    inst = Path(app.config.get("installers_dir", ""))
    files = ([f for f in inst.rglob("*") if f.is_file()]
             if inst.is_dir() else [])
    if not files:
        out("installers folder not found or empty -- backing up "
            "records only", err=True)
    data_files = ["ledger.json", "transactions.txt", "catalog.json",
                  "config.json", "mapping.json", "quarantine.json"]
    total = packed = 0
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            out(f"packing {len(files)} installer file(s)...")
            grand = sum(f.stat().st_size for f in files) or 1
            done_b, last = 0, -5
            for i, f in enumerate(files, 1):
                try:
                    z.write(f, "Installers/" + str(f.relative_to(inst)))
                    total += f.stat().st_size
                    packed += 1
                except OSError as e:
                    out(f"  could not pack {f.name}: {e}", err=True)
                done_b += f.stat().st_size
                pct = int(done_b * 100 / grand)
                if pct >= last + 5 or i == len(files):
                    last = pct
                    out(f"[dl] {pct}%  ({i}/{len(files)})")
            for name in data_files:
                f = DATA_DIR / name
                if f.exists():
                    z.write(f, "Data/" + name)
            z.writestr(
                "BACKUP_README.txt",
                "Run8 DLC Manager backup (one zip file)\n"
                f"Created: {datetime.now():%Y-%m-%d %H:%M}\n\n"
                "Installers\\  -- your DLC installer EXEs\n"
                "Data\\        -- ledger (transaction IDs), catalog, "
                "config\n\n"
                "To restore after a failure: unzip this file anywhere,\n"
                "reinstall the manager, point its Installers folder at\n"
                "the unzipped Installers\\ (or copy it back), then copy\n"
                "the Data files into the manager's data folder before\n"
                "first launch. Every transaction ID survives.\n")
        tmp.replace(zpath)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except OSError:
                pass
    sz = zpath.stat().st_size if zpath.exists() else 0
    out(f"backup complete: {packed} installer file(s) + records packed "
        f"-- {size_str(total)} in, {size_str(sz)} zip.")
    out(f"location: {zpath}")
    return 0


def cmd_restore_backup(app, args):
    """Unzip a Run8DLC_Backup.zip over the current setup. Dry-run by
    default; --yes applies. Current record files are snapshotted to a
    pre_restore_<stamp> folder first, so even a restore is reversible."""
    import zipfile
    z = Path(args.zipfile).expanduser()
    if not z.is_file():
        out(f"ERROR: {z} not found")
        return 1
    zf = zipfile.ZipFile(z)
    names = zf.namelist()
    data = [n for n in names
            if n.startswith("Data/") and not n.endswith("/")]
    inst = [n for n in names
            if n.startswith("Installers/") and not n.endswith("/")]
    tot = sum(zf.getinfo(n).file_size for n in inst)
    out(f"backup holds {len(data)} record file(s) and {len(inst)} "
        f"installer file(s) ({size_str(tot)})")
    if not args.yes:
        out("DRY RUN -- nothing restored. Re-run with --yes to apply.")
        return 0
    snap = DATA_DIR / ("pre_restore_"
                       + datetime.now().strftime("%Y%m%d_%H%M%S"))
    snap.mkdir(parents=True, exist_ok=True)
    for name in ("ledger.json", "transactions.txt", "catalog.json",
                 "config.json", "mapping.json", "quarantine.json"):
        f = DATA_DIR / name
        if f.exists():
            shutil.copy2(f, snap / name)
    out(f"current records snapshotted -> {snap}")
    for n in data:
        (DATA_DIR / Path(n).name).write_bytes(zf.read(n))
    out(f"{len(data)} record file(s) restored")
    cfg = load_json(DATA_DIR / "config.json", {})
    idir = Path(cfg.get("installers_dir")
                or app.config.get("installers_dir", ""))
    idir.mkdir(parents=True, exist_ok=True)
    done_b, last = 0, -5
    for i, n in enumerate(inst, 1):
        rel = Path(*Path(n).parts[1:])
        tgt = idir / rel
        tgt.parent.mkdir(parents=True, exist_ok=True)
        with zf.open(n) as fsrc, open(tgt, "wb") as fdst:
            shutil.copyfileobj(fsrc, fdst, 1 << 20)
        done_b += zf.getinfo(n).file_size
        pct = int(done_b * 100 / max(tot, 1))
        if pct >= last + 5 or i == len(inst):
            last = pct
            out(f"[dl] {pct}%  ({i}/{len(inst)})")
    out(f"restore complete -- {len(inst)} installer(s) -> {idir}")
    out("restart the manager to load the restored settings")
    return 0


# ------------------------------------------------------------------ doctor

def cmd_doctor(app, args):
    """Find and repair catalog/ledger damage: duplicate products created
    by 'refresh' (the store lists some items under two URL spellings),
    dead ledger rows from failed purchases, and ledger entries missing
    their installer reference."""
    plan, apply = [], args.apply

    # -- 1. catalog: fold Uncategorized dupes into their canonical product
    canonical = [p for p in app.catalog["products"]
                 if p.get("category") != "Uncategorized"]
    seen_ids = {p["id"] for p in canonical}
    can_by_sq = {}
    for p in canonical:
        can_by_sq[squash(p["name"])] = p
        for h in p.get("hints", []):
            can_by_sq.setdefault(squash(h), p)
    drops = []
    for p in [x for x in app.catalog["products"]
              if x.get("category") == "Uncategorized"]:
        target = None
        if p["id"] in seen_ids:
            target = next(c for c in canonical if c["id"] == p["id"])
        else:
            nsq = re.sub(r"^run8", "", squash(p.get("name") or ""))
            for csq, c in can_by_sq.items():
                if nsq and (nsq in csq or csq in nsq) and \
                        min(len(nsq), len(csq)) >= 6:
                    target = c
                    break
        if target:
            drops.append((p, target))
    for dupe, target in drops:
        note = ""
        if dupe.get("price") is not None and target.get("price") is None:
            note = f" (kept its ${dupe['price']:.2f} price)"
            if apply:
                target["price"] = dupe["price"]
        plan.append(f"catalog: fold duplicate '{dupe['name']}' "
                    f"[id {dupe['id']}] into '{target['name']}'{note}")
        if apply:
            app.catalog["products"].remove(dupe)

    # -- 2. ledger: drop dead rows from failed purchases
    tx_dir = Path(app.config["transactions_dir"])
    items = app.ledger.get("items", [])
    dead = []
    for it in items:
        if it.get("product_id"):
            continue
        if it.get("transaction_id") or it.get("txid"):
            continue    # a real purchase record -- never delete
        receipt = it.get("receipt")
        receipt_gone = (not receipt) or not (tx_dir / receipt).exists()
        if not it.get("exe") and receipt_gone:
            dead.append(it)
    for it in dead:
        plan.append(f"ledger: remove dead row '{it.get('name')}' "
                    f"(no product, no installer, receipt gone)")
        if apply:
            items.remove(it)

    # -- 3. ledger: backfill missing installer references
    for it in items:
        pid = it.get("product_id")
        if pid and not it.get("exe"):
            exe = find_installer_for(app, pid)
            if exe:
                plan.append(f"ledger: set installer for "
                            f"'{it.get('name')}' -> {exe.name}")
                if apply:
                    it["exe"] = exe.name
        if pid in app.by_id:
            prod = app.by_id[pid]
            if it.get("category") != prod.get("category") and apply:
                it["category"] = prod.get("category")

    if not plan:
        out("Nothing to repair -- catalog and ledger look healthy.")
        return 0
    for ln in plan:
        out(("FIXED  " if apply else "would fix  ") + ln)
    if apply:
        save_json(app.catalog_path, app.catalog)
        save_json(app.ledger_path, app.ledger)
        write_transactions_txt(app)
        out(f"\n{len(plan)} repairs applied. Rescan / re-run report to see "
            "the clean state.")
    else:
        out(f"\n{len(plan)} repairs planned. Re-run with --apply to fix.")
    return 0


def merge_purchase_records(app, pairs, source, overwrite=False):
    """pairs: (txid, pid_or_None, origin_desc, date_str). Fills blank
    transaction IDs, creates entries for owned-but-unledgered products.
    Never overwrites an existing id unless asked. Returns (rows, changed)."""
    items = app.ledger.setdefault("items", [])
    by_pid = {}
    for it in items:
        by_pid.setdefault(it.get("product_id"), []).append(it)
    rows, changed = [], 0
    for txid, pid, origin, date in pairs:
        pname = app.by_id[pid]["name"] if pid in app.by_id else "?"
        rows.append((origin, pname, txid))
        if pid is None:
            continue
        target = None
        for it in by_pid.get(pid, []):
            if not it.get("transaction_id") or overwrite:
                target = it
                break
        if target is None and by_pid.get(pid):
            continue
        if target is None:
            prod = app.by_id.get(pid)
            target = {"product_id": pid,
                      "name": prod["name"] if prod else origin,
                      "category": prod.get("category") if prod else None,
                      "exe": None, "receipt": None,
                      "transaction_id": None, "date": date,
                      "source": source}
            items.append(target)
            by_pid.setdefault(pid, []).append(target)
        if target.get("transaction_id") != txid:
            target["transaction_id"] = txid
            changed += 1
    return rows, changed


# ------------------------------------------------- document record import

def _read_eml(path):
    """Saved email receipts (Gmail: 'Download message'; Outlook: drag the
    mail to a folder). Prefers the plain-text part; falls back to
    stripped HTML."""
    import email
    from email import policy
    msg = email.message_from_bytes(Path(path).read_bytes(),
                                   policy=policy.default)
    plain, htmlpart = [], []
    for part in msg.walk():
        ct = part.get_content_type()
        if ct == "text/plain":
            try:
                plain.append(part.get_content())
            except Exception:
                pass
        elif ct == "text/html":
            try:
                htmlpart.append(part.get_content())
            except Exception:
                pass
    if plain:
        body = "\n".join(plain)
    elif htmlpart:
        body = html_to_text("\n".join(htmlpart))
    else:
        body = ""
    subj = str(msg.get("Subject") or "")
    return subj + "\n" + body


def _read_docx(path):
    import zipfile
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    xml = re.sub(r"</w:p>", "\n", xml)
    return html.unescape(re.sub(r"<[^>]+>", "", xml))


def _read_xlsx(path):
    """Rows -> lines, resolving shared strings; scientific-notation
    numbers (how Excel stores 12-digit ids) converted back to integers."""
    import zipfile
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            sx = z.read("xl/sharedStrings.xml").decode("utf-8", "replace")
            shared = [html.unescape(re.sub(r"<[^>]+>", "",
                                           m.group(1) or ""))
                      for m in re.finditer(r"<si>(.*?)</si>", sx, re.S)]
        lines = []
        for name in z.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            sheet = z.read(name).decode("utf-8", "replace")
            for row in re.finditer(r"<row[ >].*?</row>", sheet, re.S):
                vals = []
                for c in re.finditer(r"<c([^>]*)>(.*?)</c>",
                                      row.group(0), re.S):
                    attrs, body = c.group(1), c.group(2)
                    tm = re.search(r't="(\w+)"', attrs)
                    t = tm.group(1) if tm else None
                    vm = re.search(r"<v>(.*?)</v>", body, re.S)
                    v = vm.group(1) if vm else None
                    if v is None:
                        tm2 = re.search(r"<t[^>]*>(.*?)</t>", body, re.S)
                        if tm2:
                            vals.append(html.unescape(tm2.group(1)))
                        continue
                    if t == "s":
                        try:
                            vals.append(shared[int(v)])
                        except (ValueError, IndexError):
                            pass
                    else:
                        vv = v.strip()
                        if re.fullmatch(r"\d+(\.\d+)?[eE]\+?\d+", vv):
                            vv = str(int(float(vv)))
                        elif re.fullmatch(r"\d+\.0+", vv):
                            vv = vv.split(".")[0]
                        vals.append(vv)
                if vals:
                    lines.append(" | ".join(vals))
    return "\n".join(lines)


def _read_text(path):
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "utf-16", "cp1252"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return raw.decode("utf-8", "replace")


RECORD_EXTS = {".txt", ".csv", ".log", ".md", ".docx", ".xlsx", ".eml"}


def extract_records_from_file(app, path):
    """(pairs, skipped_lines) from one document: per line, find a
    transaction-id-looking number and match the rest against the catalog."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".eml":
        text = _read_eml(path)
    elif ext == ".docx":
        text = _read_docx(path)
    elif ext == ".xlsx":
        text = _read_xlsx(path)
    else:
        text = _read_text(path)
    pairs, skipped = [], []
    date = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d")
    lines = [l for l in text.splitlines()]

    def pid_of(chunk):
        if not squash(chunk):
            return None
        c = re.sub(r"(?<![\d.])\d{10,13}(?![\d.])", " ", chunk)
        c = re.sub(r"\b\d{1,4}([-/.])\d{1,2}\1\d{1,4}\b", " ", c)
        c = re.sub(r"\b\d{1,2}:\d{2}(:\d{2})?\b", " ", c)
        got, sc, m = app.match(c)
        if got and m == "hint":
            return got
        return got if got and sc >= 0.8 else None

    for idx, line in enumerate(lines):
        if not line.strip():
            continue
        ids = re.findall(r"(?<![\d.])(\d{10,13})(?![\d.])", line)
        if not ids:
            labeled = extract_txid_from_ocr(line)
            ids = [labeled] if labeled else []
        if not ids:
            continue
        pid = pid_of(line)
        if not pid:
            # receipts put name and id on separate lines: look nearby
            for j in list(range(idx - 1, max(-1, idx - 6), -1)) \
                    + [idx + 1, idx + 2]:
                if 0 <= j < len(lines) and lines[j].strip():
                    pid = pid_of(lines[j])
                    if pid:
                        break
        if pid:
            pairs.append((ids[0], pid, f"{path.name}: {line.strip()[:44]}",
                          date))
        else:
            skipped.append((path.name, line.strip()[:70], ids[0]))

    # a whole email with one id but a name we couldn't place line-wise
    if not pairs and ext == ".eml":
        txid = extract_txid_from_ocr(text)
        if not txid:
            m2 = re.search(r"(?<![\d.])(\d{10,13})(?![\d.])", text)
            txid = m2.group(1) if m2 else None
        if txid:
            pid = pid_of(text[:600])
            if pid:
                pairs.append((txid, pid, f"{path.name} (whole email)",
                              date))
                skipped = [s for s in skipped if s[0] != path.name]
    return pairs, skipped


def cmd_import_records(app, args):
    targets = []
    for p in args.paths:
        pp = Path(p)
        if pp.is_dir():
            targets += sorted(x for x in pp.iterdir()
                              if x.suffix.lower() in RECORD_EXTS)
        elif pp.suffix.lower() in RECORD_EXTS:
            targets.append(pp)
        elif pp.suffix.lower() in (".png", ".jpg", ".jpeg", ".bmp"):
            out(f"  {pp.name}: images go through 'ocr-receipts' (or the "
                "Import button, which handles both)")
        else:
            out(f"  {pp.name}: unsupported type ({pp.suffix}) -- convert "
                "old .doc/.xls to .docx/.xlsx first")
    if not targets:
        out("No importable files. Supported: "
            + " ".join(sorted(RECORD_EXTS)))
        return 1
    all_pairs, all_skipped = [], []
    for t in targets:
        try:
            pairs, skipped = extract_records_from_file(app, t)
            out(f"  {t.name}: {len(pairs)} record(s) recognized")
            all_pairs += pairs
            all_skipped += skipped
        except Exception as e:
            out(f"  {t.name}: could not read ({e})")
    rows, changed = merge_purchase_records(app, all_pairs, source="import",
                                           overwrite=args.overwrite)
    out("")
    for origin, pname, txid in rows:
        out(f"  {pname:<42} {txid}   ({origin})")
    if all_skipped:
        out("\nLines with a transaction-id-looking number but no "
            "recognizable product:")
        for fn, line, txid in all_skipped[:15]:
            out(f"  {fn}: {line}")
        out("  fix these with:  add --name \"<product>\" --txid <id>")
    save_json(app.ledger_path, app.ledger)
    write_transactions_txt(app)
    out(f"\n{len(rows)} records read, {changed} ledger entries updated, "
        f"{len(all_skipped)} lines skipped. transactions.txt regenerated.")
    return 0



# ------------------------------------------------------------ product media

NOISE_IMG = re.compile(r"topheader|store_hdr|line_[tbm]|banner|logo|paypal|"
                       r"visa|mastercard|secure|spacer|bullet|arrow|"
                       r"buyitnow|buy_|\.gif$", re.I)


def _imgs_in(html_chunk, base_url):
    out_list = []
    for m in re.finditer(r"""<img[^>]+src\s*=\s*["']?([^"'\s>]+)""",
                         html_chunk, re.I):
        u = m.group(1)
        if NOISE_IMG.search(u):
            continue
        out_list.append(urllib.parse.urljoin(base_url, u))
    return out_list


def _match_pid(app, text):
    """Hint-match with OCR-tolerant retries: raw, pack-number collapsed,
    and II->11 (letter/digit confusion)."""
    for t in (text,
              re.sub(r"\bpacks?\b", " ", text, flags=re.I),
              text.replace("II", "11").replace("Il", "11")):
        got, _sc, method = app.match(t)
        if got and method == "hint":
            return got
    return None


def assign_images_by_ocr(app, texts, prefer=None):
    """texts: {candidate_filename: ocr_text}. The store prints the product
    name on each image, so the images identify themselves. Returns
    ({pid: filename}, [unmatched filenames]). Shortest matching text wins
    (title banners beat long captions)."""
    prefer = prefer or {}
    by_pid = {}
    unmatched = []
    for fname, txt in texts.items():
        t = (txt or "").strip()
        pid = _match_pid(app, t) if t else None
        if pid:
            by_pid.setdefault(pid, []).append((len(t), fname))
        else:
            unmatched.append(fname)
    chosen = {}
    for pid, cands in by_pid.items():
        if pid in prefer:
            continue
        cands.sort()
        chosen[pid] = cands[0][1]
    return chosen, unmatched


def cmd_media(app, args):
    """Build the product image library: download every candidate image
    from the store's pages, OCR the product names printed on them, and
    assign each image to its product by what it says about itself."""
    media = DATA_DIR / "media"
    cand_dir = media / "cand"
    srcdir = media / "src"
    for d in (cand_dir, srcdir):
        d.mkdir(parents=True, exist_ok=True)
    mmap = load_json(media / "media_map.json", {})
    legacy = mmap and mmap.pop("_v", None) != 5
    if args.force or legacy:
        if legacy:
            out("image map from an older version detected -- rebuilding "
                "assignments from OCR evidence")
        mmap = {k: v for k, v in mmap.items()
                if str(v).startswith("manual:")}
        for f in srcdir.glob("*"):
            f.unlink(missing_ok=True)

    # 1. gather every candidate image URL from the catalog + route pages
    page_urls = list(app.config["catalog_pages"])
    page_urls += sorted({p["info_url"] for p in app.products
                         if p.get("info_url")})
    cand_urls, seen = [], set()
    out(f"scanning {len(page_urls)} store pages for images...")
    for u in page_urls:
        try:
            for img in _imgs_in(fetch(u, timeout=15), u):
                if img not in seen:
                    seen.add(img)
                    cand_urls.append(img)
            time.sleep(0.3)
        except Exception as e:
            out(f"  {u}: {e}")
    out(f"  {len(cand_urls)} candidate images found")

    # 2. download candidates (cached by URL hash)
    import hashlib
    cmap = load_json(media / "cand_map.json", {})
    url_to_file = {v: k for k, v in cmap.items()}
    todo = [u for u in cand_urls if u not in url_to_file
            or not (cand_dir / url_to_file[u]).exists()]
    out(f"downloading {len(todo)} new image(s)...")
    for i, u in enumerate(todo, 1):
        out(f"[dl] {i * 100 // max(len(todo), 1)}%")
        try:
            ext = Path(urllib.parse.urlparse(u).path).suffix or ".jpg"
            name = "cand_" + hashlib.md5(u.encode()).hexdigest()[:10] + ext
            (cand_dir / name).write_bytes(fetch(u, timeout=15, binary=True))
            cmap[name] = u
        except Exception as e:
            out(f"  {u.rsplit('/', 1)[-1]}: {e}")
    save_json(media / "cand_map.json", cmap)

    if sys.platform != "win32":
        out("(image OCR + conversion use Windows -- run this on your PC "
            "to finish)")
        return 0

    # 3. OCR the candidates so they identify themselves
    ps1 = DATA_DIR / "ocr_receipts.ps1"
    if not ps1.exists():
        ensure_assets()
    tmp = media / "cand_ocr.json"
    ocr = load_json(tmp, {})
    fresh = [f.name for f in cand_dir.iterdir()
             if f.is_file() and f.name not in ocr]
    if fresh:
        out(f"reading the product names printed on {len(fresh)} new "
            "image(s)...")
        tmp2 = media / "cand_ocr_new.json"
        r = subprocess.run(["powershell.exe", "-NoProfile",
                            "-ExecutionPolicy", "Bypass", "-File", str(ps1),
                            "-InDir", str(cand_dir), "-OutJson", str(tmp2),
                            "-Files"] + fresh, capture_output=True,
                           text=True, timeout=1200)
        if r.returncode != 0:
            out((r.stderr or "OCR failed").strip()[:300])
            return 1
        ocr.update(load_json(tmp2, {}))
        tmp2.unlink(missing_ok=True)
        save_json(tmp, ocr)
    else:
        out("using cached OCR results for all candidate images")
    texts = {k: v.get("text", "") for k, v in ocr.items() if v.get("ok")}
    manual = {k for k, v in mmap.items() if str(v).startswith("manual:")}
    chosen, unmatched = assign_images_by_ocr(app, texts, prefer=manual)

    # banners OCR couldn't read: let the image's own URL filename speak
    claimed = set(chosen.values())
    prods_all = {p["id"] for p in app.products if p.get("url")}
    for fname, u in cmap.items():
        if fname in claimed:
            continue
        stem = Path(urllib.parse.urlparse(u).path).stem
        pid = _match_pid(app, stem.replace("_", " ").replace("-", " "))
        if pid and pid in prods_all and pid not in chosen \
                and pid not in manual:
            chosen[pid] = fname
            claimed.add(fname)
    # the base sim and updates have no product banner; use the site art
    DEFAULT_ART = {
        "base_v3": "https://www.run8studios.com/images/r83_index_hdr1.jpg",
        "updates":
            "https://www.run8studios.com/images/run8_route_page_hdr1.jpg",
    }
    url_to_cand = {v: k for k, v in cmap.items()}
    for pid, u in DEFAULT_ART.items():
        if pid not in chosen and pid not in mmap and u in url_to_cand:
            chosen[pid] = url_to_cand[u]

    # 4. stage chosen images per product and convert
    assigned = 0
    for pid, fname in chosen.items():
        srcf0 = list(srcdir.glob(pid + ".*"))
        if pid in mmap and srcf0 and not args.force:
            continue
        srcf = cand_dir / fname
        if not srcf.exists():
            continue
        shutil.copy2(srcf, srcdir / (pid + srcf.suffix))
        mmap[pid] = cmap.get(fname, fname)
        assigned += 1
    # manual overrides: download directly
    for pid, v in list(mmap.items()):
        if str(v).startswith("manual:") and not (srcdir / (pid + ".jpg")).exists():
            try:
                u = str(v)[7:]
                (srcdir / (pid + ".jpg")).write_bytes(
                    fetch(u, timeout=15, binary=True))
            except Exception as e:
                out(f"  manual {pid}: {e}")
    mmap["_v"] = 5
    save_json(media / "media_map.json", mmap)

    ps = DATA_DIR / "img_convert.ps1"
    ps.write_text(IMG_PS1, encoding="utf-8")
    out("converting to display-ready PNGs...")
    r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy",
                        "Bypass", "-File", str(ps), "-SrcDir", str(srcdir),
                        "-OutDir", str(media),
                        "-NoCrop", "base_v3,updates"],
                       capture_output=True,
                       text=True, timeout=600)
    for ln in (r.stdout or "").splitlines()[-3:]:
        out("  " + ln)

    have = {f.stem for f in media.glob("*_t.png")}
    missing = [p["name"] for p in app.products
               if p.get("url") and p["id"] not in have]
    out(f"\n{assigned} newly assigned by OCR; "
        f"{len(have)} of {sum(1 for p in app.products if p.get('url'))} "
        "products now have images.")
    if missing:
        out("still without an image: " + ", ".join(missing[:12])
            + ("..." if len(missing) > 12 else ""))
        out("fix any by adding  \"pid\": \"manual:<image url>\"  to "
            "media_map.json and re-running media")
    return 0


# ------------------------------------------------------------- receipt OCR

OCR_CACHE = DATA_DIR / "ocr_cache.json"


def extract_txid_from_ocr(text):
    """Pull a transaction id out of noisy OCR text. Real 3DTS receipts:
    'Transaction ID: 62462829837' (OCR often reads ID as 10/I0/lD) and a
    'DL LINK: ...download.php?transid=62462829837'. IDs are numeric,
    ~10-13 digits. The URL sometimes truncates mid-OCR, so gather every
    candidate and keep the longest."""
    t = text.replace("\u00a0", " ")
    cands = []
    for m in re.finditer(r"transid\s*=\s*(\d{6,16})", t, re.I):
        cands.append(m.group(1))
    for m in re.finditer(r"Trans\w*\s*[1Il|]?[DdOo0]\s*[:.#]?\s*"
                         r"((?:\d ?){8,16}\d)", t):
        cands.append(re.sub(r"\s", "", m.group(1)))
    m = re.search(r"Trans\w*\s*[1Il|]?[Dd]\s*[:.#]?\s*([A-Za-z0-9]{10,24})", t)
    if m and not m.group(1).isdigit():
        cands.append(m.group(1).upper())
    m = re.search(r"\b([A-Z0-9]{17})\b", t)
    if m and not m.group(1).isdigit():
        cands.append(m.group(1))
    return max(cands, key=len) if cands else None


def cmd_ocr_receipts(app, args):
    tx_dir = Path(app.config["transactions_dir"])
    if not tx_dir.is_dir():
        out(f"Transactions folder not found: {tx_dir}")
        return 1
    images = sorted(f for f in tx_dir.iterdir() if f.is_file()
                    and f.suffix.lower() in (".png", ".jpg", ".jpeg", ".bmp"))
    if not images:
        out("No receipt images found.")
        return 1

    cache = load_json(OCR_CACHE, {})
    leftover = DATA_DIR / "ocr_output.json"
    if leftover.exists():
        prev = load_json(leftover, {})
        by_name = {f.name: f for f in images}
        for name, rec in prev.items():
            if name in by_name and rec.get("ok"):
                rec["mtime"] = int(by_name[name].stat().st_mtime)
                cache[name] = rec
        if prev:
            out(f"absorbed {len(prev)} results from a previous OCR run")
        save_json(OCR_CACHE, cache)
        leftover.unlink()
    todo = [f for f in images
            if args.force or f.name not in cache
            or not cache[f.name].get("ok")
            or int(cache[f.name].get("mtime", 0)) != int(f.stat().st_mtime)]

    if todo:
        if sys.platform != "win32":
            out(f"{len(todo)} image(s) need OCR, which uses Windows' "
                "built-in engine -- run this on your PC.")
            return 1
        ps1 = DATA_DIR / "ocr_receipts.ps1"
        tmp = DATA_DIR / "ocr_output.json"
        out(f"OCR-ing {len(todo)} receipt image(s) with Windows OCR...")
        cmd = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
               "-File", str(ps1), "-InDir", str(tx_dir),
               "-OutJson", str(tmp)]
        if len(todo) < len(images):
            cmd += ["-Files"] + [f.name for f in todo]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True,
                               timeout=1200)
        except subprocess.TimeoutExpired:
            out("OCR timed out.")
            return 1
        for ln in (r.stdout or "").splitlines():
            out("  " + ln)
        if r.returncode != 0:
            out((r.stderr or "").strip())
            out("OCR failed -- see message above.")
            return 1
        fresh = load_json(tmp, {})
        # normalize mtimes to our own stat so cache comparisons hold
        for f in todo:
            if f.name in fresh:
                fresh[f.name]["mtime"] = int(f.stat().st_mtime)
        cache.update(fresh)
        save_json(OCR_CACHE, cache)
        tmp.unlink(missing_ok=True)

    # parse + match + merge
    pairs, unresolved = [], []
    for f in images:
        rec = cache.get(f.name) or {}
        txid = extract_txid_from_ocr(rec.get("text", "")) if rec.get("ok") \
            else None
        if not txid:
            unresolved.append((f.name, "no transaction id readable"))
            continue
        pid, _sc, _m = app.match(f.name)
        pairs.append((txid, pid, f.name,
                      datetime.fromtimestamp(f.stat().st_mtime
                                             ).strftime("%Y-%m-%d")))
    rows, changed = merge_purchase_records(app, pairs, source="ocr",
                                           overwrite=args.overwrite)

    out("")
    for name, pname, txid in rows:
        out(f"  {name:<38} {pname:<38} {txid}")
    if unresolved:
        out("")
        out("Could not read a transaction id from:")
        for name, why in unresolved:
            out(f"  {name}  ({why})")
        out("  (run with --show-text \"<filename>\" to see the raw OCR text,")
        out("   or send ocr_cache.json to Claude to tune the parser)")
    save_json(app.ledger_path, app.ledger)
    write_transactions_txt(app)
    out(f"\n{len(rows)} receipts read, {changed} ledger entries updated, "
        f"{len(unresolved)} unresolved. transactions.txt regenerated.")
    return 0


def cmd_ocr_show(app, args):
    cache = load_json(OCR_CACHE, {})
    rec = cache.get(args.show_text)
    if not rec:
        out(f"no cached OCR for '{args.show_text}' -- run ocr-receipts first")
        return 1
    out(rec.get("text") or f"(failed: {rec.get('error')})")
    return 0


# ----------------------------------------------------------------- refresh

SLUG_RE = re.compile(r"store_[A-Za-z0-9_\-]+\.php")
PRICE_RE = re.compile(r"price\s*:?\s*\$\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)", re.I)


def parse_store_page(html_text):
    """Store pages title like '=== Run8 MP15 Pack 3 - RUN8 - 3DTS Secure...==='
    and contain 'price: $20.00 (USD)' in the body."""
    name, price = None, None
    m = re.search(r"<title>(.*?)</title>", html_text, re.I | re.S)
    if m:
        t = re.sub(r"\s+", " ", m.group(1)).replace("=", "").strip()
        t = re.split(r"\s*-\s*RUN8\b", t, flags=re.I)[0]
        name = t.strip(" -") or None
    m = PRICE_RE.search(html_to_text(html_text))
    if m:
        price = float(m.group(1).replace(",", ""))
    return name, price


def cmd_refresh(app, args):
    known_slugs = set()
    for p in app.products:
        for u in [p.get("url")] + list(p.get("alt_urls", [])):
            if u:
                known_slugs.add(
                    Path(urllib.parse.urlparse(u).path).name.lower())

    found = {}
    for page_url in app.config["catalog_pages"]:
        out(f"Scanning {page_url} ...")
        try:
            h = fetch(page_url)
        except Exception as e:
            out(f"  ERROR fetching page: {e}")
            continue
        if args.dump_html:
            dump = DATA_DIR / ("dump_" + safe_name(Path(page_url).name) + ".html")
            dump.write_text(h, encoding="utf-8")
            out(f"  raw HTML -> {dump}")
        for slug in set(SLUG_RE.findall(h)):
            found[slug.lower()] = urllib.parse.urljoin(
                app.config["store_base"], slug)
        time.sleep(0.5)

    new_slugs = sorted(s for s in found if s not in known_slugs)
    added = 0
    if new_slugs:
        out(f"\n{len(new_slugs)} NEW products found in the store:")
        for slug in new_slugs:
            url = found[slug]
            name, price = None, None
            try:
                name, price = parse_store_page(fetch(url))
            except Exception as e:
                out(f"  (could not read {url}: {e})")
            time.sleep(0.5)
            core = re.sub(r"^store_(run8[_\-]?)?", "", Path(slug).stem)
            pid = squash(core) or slug
            # same product under an alternate URL spelling? update, don't dupe
            existing = app.by_id.get(pid)
            if existing is None and name:
                got, sc, _m = app.match(name)
                if got and sc >= 0.9:
                    existing = app.by_id[got]
            if existing is None:
                got, sc, _m = app.match(core)
                if got and sc >= 0.9:
                    existing = app.by_id[got]
            if existing is not None:
                if price is not None and existing.get("price") is None:
                    existing["price"] = price
                alts = existing.setdefault("alt_urls", [])
                if url not in alts:
                    alts.append(url)
                out(f"  = {existing['name']}: alternate store URL "
                    f"({slug}) -- remembered, so it won't count as "
                    "new again")
                continue
            entry = {"id": pid, "name": name or core, "category": "Uncategorized",
                     "price": price, "url": url, "hints": [squash(core)]}
            app.products.append(entry)
            app.by_id[pid] = entry
            added += 1
            out(f"  + {entry['name']}  {price_str(entry)}  ({url})")
        if added:
            out("\nNew items were added with category 'Uncategorized' - "
                "edit catalog.json to set Route/Locomotive/etc. and add "
                "hints.")
        else:
            out("\n(all were alternate links to products already in "
                "the catalog -- nothing added)")
    else:
        out("\nNo new products since the bundled catalog.")

    if args.prices:
        out("\nUpdating prices from individual store pages "
            f"({sum(1 for p in app.products if p.get('url'))} pages, ~0.5s each)...")
        total_n = sum(1 for x in app.products if x.get("url"))
        done_n = 0
        for p in app.products:
            if not p.get("url") or "3dts-onlinestore" not in p["url"]:
                continue
            done_n += 1
            out(f"[dl] {done_n * 100 // max(total_n, 1)}%  {p['name']}")
            try:
                _n, price = parse_store_page(fetch(p["url"], timeout=10))
                if price is not None and price != p.get("price"):
                    out(f"  {p['name']:<45} ${price:.2f}")
                    p["price"] = price
            except Exception as e:
                out(f"  {p['name']}: fetch failed ({e})")
            time.sleep(0.5)

    save_json(app.catalog_path, {
        "catalog_date": datetime.now().date().isoformat(),
        "source_pages": app.config["catalog_pages"],
        "products": app.products,
    })
    out(f"\nCatalog saved ({len(app.products)} entries) -> {app.catalog_path}")


# ---------------------------------------------------------------- snapshot

def cmd_snapshot(app, args):
    lines = [f"run8dlc snapshot {datetime.now().isoformat()}",
             f"tool version {VERSION}", ""]

    root = Path(app.config["run8_install"])
    lines.append(f"== Run8 install: {root} ==")
    if root.is_dir():
        root_depth = len(root.parts)
        for dirpath, dirnames, files in os.walk(root):
            depth = len(Path(dirpath).parts) - root_depth
            if depth > 4:
                dirnames[:] = []
                continue
            rel = Path(dirpath).relative_to(root)
            lines.append(f"{'  ' * depth}{rel if str(rel) != '.' else '(root)'}/"
                         f"  [{len(files)} files]")
    else:
        lines.append("  (directory not found)")

    # full filename dumps of the shared equipment folders -- this is what
    # lets family-token detection get refined against reality
    for sub in (("Content", "V3RailVehicles", "Body"),
                ("Content", "V3RailVehicles", "Trucks")):
        d = root.joinpath(*sub)
        lines.append(f"\n== {Path(*sub)} files ==")
        if d.is_dir():
            for f in sorted(d.iterdir()):
                if f.is_file():
                    lines.append(f"  {f.name}")
        else:
            lines.append("  (directory not found)")

    for label, key in (("Installers", "installers_dir"),
                       ("Transactions", "transactions_dir")):
        d = Path(app.config[key])
        lines.append(f"\n== {label}: {d} ==")
        if d.is_dir():
            for f in sorted(d.iterdir()):
                if f.is_file():
                    lines.append(f"  {f.name}")
        else:
            lines.append("  (directory not found)")

    dest = DATA_DIR / "run8_snapshot.txt"
    dest.write_text("\n".join(lines), encoding="utf-8")
    out(f"Snapshot written -> {dest}")
    out("Send this file back to Claude to tune in-game DLC detection.")


# -------------------------------------------------------------------- main

def cmd_installed(app, args):
    state, _un = app.build_state(include_game_scan=True)
    qp = quarantined_pids()
    rows = []
    for p in app.products:
        if p.get("category") == "Base" and not args.all:
            continue
        st = product_status(p["id"], state, qp)
        ev = state.get(p["id"], {})
        detail = ""
        if st == "installed":
            detail = (ev["game"] or ev.get("game_files", ["?"]))[0]
        elif st in ("owned", "quarantined") and ev.get("installers"):
            detail = ev["installers"][-1].name
        rows.append((p.get("category", ""), p["name"], st, detail))
    rows.sort()
    ICON = {"installed": "[#]", "owned": "[o]",
            "quarantined": "[q]", "missing": "[ ]"}
    counts = {}
    cat = None
    for c, name, st, detail in rows:
        if c != cat:
            out(f"\n{c}")
            cat = c
        counts[st] = counts.get(st, 0) + 1
        out(f"  {ICON[st]} {name:<44} {st:<11} {detail}")
    out("")
    out("  [#] installed (detected in game)   [o] owned (installer on file)")
    out("  [q] quarantined by 'uninstall'     [ ] not owned")
    out(f"\n  {counts.get('installed', 0)} installed, "
        f"{counts.get('owned', 0)} owned-not-detected, "
        f"{counts.get('quarantined', 0)} quarantined, "
        f"{counts.get('missing', 0)} not owned")
    out("  note: equipment packs usually show 'owned' -- in-game detection "
        "keys off route folder names until snapshot tuning")


def _resolve_or_complain(app, query):
    pid, cands = resolve_product(app, query)
    if not pid:
        out(f"couldn't identify a product from '{query}'")
        for c in cands:
            out(f"  maybe: {c}")
        return None
    return pid


def cmd_reinstall(app, args):
    pid = _resolve_or_complain(app, args.query)
    if not pid:
        return 1
    exe = find_installer_for(app, pid)
    name = app.by_id[pid]["name"]
    if not exe:
        out(f"no installer EXE found in Installers\\ for {name}")
        out("  (download it again via 'add', or drop the EXE in that folder)")
        return 1
    out(f"{name}\n  installer: {exe}")
    if args.dry_run:
        out("  dry run -- not launching")
        return 0
    launched = launch(exe, cwd=exe.parent)
    out("  launched." if launched else "")
    return 0


def cmd_uninstall(app, args):
    pid = _resolve_or_complain(app, args.query)
    if not pid:
        return 1
    ok, lines = uninstall_product(app, pid, apply=args.apply)
    for ln in lines:
        out(ln)
    return 0 if ok else 1


def cmd_restore(app, args):
    pid = _resolve_or_complain(app, args.query)
    if not pid:
        return 1
    ok, lines = restore_product(app, pid)
    for ln in lines:
        out(ln)
    return 0 if ok else 1


def cmd_updater(app, args):
    exe = updater_path(app)
    if not exe.is_file():
        out(f"updater not found: {exe}")
        out("  set \"updater_exe\" in config.json if it lives elsewhere")
        return 1
    out(f"launching {exe}")
    launch(exe, cwd=exe.parent)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog="run8dlc",
                                 description="Run8 DLC library manager")
    ap.add_argument("--config", help="path to config.json "
                    "(default: next to script)")
    sub = ap.add_subparsers(dest="cmd")

    rp = sub.add_parser("report", help="diff owned vs store (default)")
    rp.add_argument("--open", action="store_true",
                    help="open report.html when done")
    rp.add_argument("--verbose", action="store_true",
                    help="print unmatched files to console")
    rp.add_argument("--no-game-scan", action="store_true",
                    help="skip scanning the Run8 install directory")

    mg = sub.add_parser("migrate", help="build ledger from existing folders")
    mg.add_argument("--organize", action="store_true",
                    help="rename receipts to '<Product Name>.png'")
    mg.add_argument("--apply", action="store_true",
                    help="actually perform --organize renames "
                    "(otherwise dry run)")

    ad = sub.add_parser("add", help="record a new purchase")
    ad.add_argument("url", nargs="?", help="the receipt/download page URL "
                    "(copy it from the browser right after purchase)")
    ad.add_argument("--exe", help="path to an already-downloaded installer")
    ad.add_argument("--receipt", help="path to a manual receipt screenshot")
    ad.add_argument("--name", help="product name override")
    ad.add_argument("--txid", help="transaction ID override")
    ad.add_argument("--link", help="direct download link (skips the "
                    "receipt page)")
    ad.add_argument("--install", action="store_true",
                    help="launch the installer after saving")
    ad.add_argument("--dump", action="store_true",
                    help="save raw receipt-page HTML for debugging")

    rf = sub.add_parser("refresh", help="re-scan store for new products")
    rf.add_argument("--prices", action="store_true",
                    help="also fetch every product page to update prices")
    rf.add_argument("--dump-html", action="store_true",
                    help="save raw catalog-page HTML for debugging")

    sub.add_parser("snapshot", help="dump install-dir listing for diagnostics")

    ins = sub.add_parser("installed", help="list every product's install status")
    ins.add_argument("--all", action="store_true",
                     help="include base-sim / updater rows")

    ri = sub.add_parser("reinstall", help="launch a product's installer EXE")
    ri.add_argument("query", help="product name / id fragment")
    ri.add_argument("--dry-run", action="store_true",
                    help="show which EXE would run, don't launch")

    un = sub.add_parser("uninstall",
                        help="quarantine a product's in-game folders (reversible)")
    un.add_argument("query", help="product name / id fragment")
    un.add_argument("--apply", action="store_true",
                    help="actually move folders (default: dry run)")

    rs = sub.add_parser("restore", help="undo an uninstall")
    rs.add_argument("query", help="product name / id fragment")

    sub.add_parser("updater", help="launch Run8_Updater.exe")

    sub.add_parser("transactions",
                   help="print the purchase record (also transactions.txt)")

    dr = sub.add_parser("doctor", help="find & repair catalog/ledger damage "
                        "(duplicate products, dead purchase rows, missing "
                        "installer references)")
    dr.add_argument("--apply", action="store_true",
                    help="apply the repairs (default: show the plan)")

    md = sub.add_parser("media", help="download one image per product and convert for in-app display")
    md.add_argument("--force", action="store_true",
                    help="re-locate and re-download everything")

    pb = sub.add_parser("backup", help="copy installers + transaction "
                        "records to another drive")
    pb.add_argument("dest", nargs="?", default=None,
                    help="destination folder (default: the configured "
                         "Backups folder)")

    rb = sub.add_parser("restore-backup",
                        help="restore records + installers from a "
                             "Run8DLC_Backup.zip (dry run unless --yes)")
    rb.add_argument("zipfile", help="path to Run8DLC_Backup.zip")
    rb.add_argument("--yes", action="store_true",
                    help="actually restore (otherwise dry run)")

    pq = sub.add_parser("purge-quarantine",
                        help="PERMANENTLY delete everything in the "
                             "quarantine folder")
    pq.add_argument("--yes", action="store_true",
                    help="confirm permanent deletion")

    sub.add_parser("emit-assets", help="write bundled catalog/icon/scripts "
                   "to disk (used by the EXE build)")

    ir = sub.add_parser("import-records",
                        help="import transaction IDs from txt/csv/docx/xlsx "
                             "purchase records")
    ir.add_argument("paths", nargs="+", help="files or folders to scan")
    ir.add_argument("--overwrite", action="store_true",
                    help="replace transaction IDs already in the ledger")

    oc = sub.add_parser("ocr-receipts",
                        help="read transaction IDs out of receipt PNGs "
                             "(Windows built-in OCR) into the ledger")
    oc.add_argument("--force", action="store_true",
                    help="re-OCR everything, ignoring the cache")
    oc.add_argument("--overwrite", action="store_true",
                    help="replace transaction IDs already in the ledger")
    oc.add_argument("--show-text", metavar="FILENAME", default=None,
                    help="print the raw OCR text for one receipt and exit")

    args = ap.parse_args(argv)
    app = App(args.config)

    if args.cmd is None:
        args = ap.parse_args((argv or sys.argv[1:]) + ["report"])
        args.cmd = "report"

    if args.cmd == "report":
        return cmd_report(app, args)
    if args.cmd == "migrate":
        return cmd_migrate(app, args)
    if args.cmd == "add":
        if not (args.url or args.exe or args.txid
                or getattr(args, "link", None)):
            out("add: provide a receipt-page URL, --exe <file>, "
                "--txid <id>, or --link <download url>")
            return 1
        return cmd_add(app, args)
    if args.cmd == "refresh":
        return cmd_refresh(app, args)
    if args.cmd == "snapshot":
        return cmd_snapshot(app, args)
    if args.cmd == "installed":
        return cmd_installed(app, args)
    if args.cmd == "reinstall":
        return cmd_reinstall(app, args)
    if args.cmd == "uninstall":
        return cmd_uninstall(app, args)
    if args.cmd == "restore":
        return cmd_restore(app, args)
    if args.cmd == "updater":
        return cmd_updater(app, args)
    if args.cmd == "transactions":
        return cmd_transactions(app, args)
    if args.cmd == "doctor":
        return cmd_doctor(app, args)
    if args.cmd == "backup":
        return cmd_backup(app, args)
    if args.cmd == "restore-backup":
        return cmd_restore_backup(app, args)
    if args.cmd == "purge-quarantine":
        return cmd_purge_quarantine(app, args)
    if args.cmd == "media":
        return cmd_media(app, args)

    if args.cmd == "import-records":
        return cmd_import_records(app, args)
    if args.cmd == "emit-assets":
        ensure_assets()
        out(f"assets written to {DATA_DIR}")
        return 0
    if args.cmd == "ocr-receipts":
        if args.show_text:
            return cmd_ocr_show(app, args)
        return cmd_ocr_receipts(app, args)




def restart_app():
    """Relaunch this program (used when the color theme changes)."""
    if getattr(sys, "frozen", False):
        os.execl(sys.executable, sys.executable)
    else:
        os.execl(sys.executable, sys.executable,
                 str(Path(__file__).resolve()))


def create_shortcuts():
    """Desktop + Start Menu shortcuts pointing at this app (works for both
    the .pyw and the frozen EXE)."""
    if sys.platform != "win32":
        return False, "Windows only."
    if getattr(sys, "frozen", False):
        target, sc_args = sys.executable, ""
    else:
        pyw = Path(sys.executable).with_name("pythonw.exe")
        target = str(pyw if pyw.exists() else sys.executable)
        sc_args = f'\"{Path(__file__).resolve()}\"'
    icon = DATA_DIR / "run8dlc.ico"
    ps = f"""$ws = New-Object -ComObject WScript.Shell
$dirs = @([Environment]::GetFolderPath('Desktop'), "$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs")
foreach ($d in $dirs) {{
    $s = $ws.CreateShortcut((Join-Path $d 'Run8 DLC Manager.lnk'))
    $s.TargetPath = '{target}'
    $s.Arguments = '{sc_args}'
    $s.WorkingDirectory = '{DATA_DIR}'
    $s.IconLocation = '{icon},0'
    $s.Save()
}}"""
    p1 = DATA_DIR / "_mk_shortcut.ps1"
    p1.write_text(ps, encoding="utf-8")
    r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy",
                        "Bypass", "-File", str(p1)], capture_output=True,
                       text=True)
    p1.unlink(missing_ok=True)
    if r.returncode == 0:
        return True, "Shortcuts created on the Desktop and Start Menu."
    return False, (r.stderr or "PowerShell refused.").strip()[:400]


import queue
import threading
import webbrowser

try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    import tkinter.font as tkfont
    HAVE_TK = True
except ImportError:
    HAVE_TK = False

if HAVE_TK:
    # ------------------------------------------------------------------ theme
    # Fixed, legibility-tested palettes. Order: BG, PANEL, PANEL2, FIELD,
    # FG, DIM, ACCENT, ACC_FG (text on accent), OWNEDC, MISSC, QUARC, SEL
    PALETTES = {
        "Boxcar Slate":
            ("#14171c", "#1b2129", "#242c37", "#10131a",
             "#d7dde6", "#8b96a5", "#37b6a7", "#10151d",
             "#6fa7e0", "#c9705f", "#a98ae0", "#2a3646",
             "#6fa7e0"),
        "Amtrak Phase III":
            ("#121318", "#181c22", "#20252e", "#0d1015",
             "#d7dbe0", "#8e929c", "#d5212e", "#f4f6f9",
             "#7fa8d8", "#e8b06f", "#a98ae0", "#292f3a",
             "#3d6fd6"),
        "UP Armour Yellow":
            ("#1a1917", "#222422", "#2d2e2d", "#141414",
             "#dcdee0", "#95989c", "#ffb612", "#12141a",
             "#b1b3b3", "#e07a6f", "#a98ae0", "#393c38",
             "#da291c"),
        "ATSF Warbonnet":
            ("#161417", "#1e1d22", "#27272d", "#111115",
             "#dadbe0", "#92939c", "#c8102e", "#f4f6f9",
             "#c7c9c7", "#ffc72c", "#a98ae0", "#323239",
             "#c7c9c7"),
        "BNSF Heritage II":
            ("#151415", "#1c1c20", "#25262a", "#101013",
             "#d9dbdf", "#91939a", "#ff6720", "#12141a",
             "#9aa0a4", "#d0707f", "#b48fe0", "#2f3035",
             "#ffcd00"),
        "SP Daylight":
            ("#171517", "#1f1e21", "#29282c", "#121114",
             "#dadce0", "#93949b", "#e86a1f", "#12141a",
             "#e8c06f", "#d0708c", "#a98ae0", "#343437",
             "#c8102e"),
        "Conrail Blue":
            ("#12161b", "#191f26", "#212932", "#0e1217",
             "#d8dce2", "#8f949f", "#0079c1", "#f4f6f9",
             "#9fb8cc", "#d0705f", "#a98ae0", "#2a343f",
             "#e8edf2"),
        "CSX Blue & Gold":
            ("#11161c", "#181f28", "#202934", "#0d1218",
             "#d7dce3", "#8e94a0", "#fdb813", "#12141a",
             "#5e9bd0", "#d0705f", "#a98ae0", "#283443",
             "#2d7dd2"),
        "Rio Grande Gold":
            ("#161616", "#1e1f20", "#27292b", "#111213",
             "#dadcdf", "#92949b", "#ffb81c", "#12141a",
             "#c8b49a", "#d0705f", "#a98ae0", "#323436",
             "#e8e2d0"),
        "Chessie Yellow":
            ("#12161b", "#191f26", "#212933", "#0e1217",
             "#d8dce2", "#8e949f", "#ffc425", "#12141a",
             "#7fa8d8", "#d0708c", "#a98ae0", "#2a3440",
             "#e8542f"),
        "BN Cascade Green":
            ("#111617", "#181f22", "#20292d", "#0d1214",
             "#dbe6e0", "#8e949c", "#008249", "#f4f6f9",
             "#9fc7b4", "#d0705f", "#a98ae0", "#283438",
             "#e8f0ea"),
        "High Contrast":
            ("#000000", "#0d0d0d", "#1a1a1a", "#000000",
             "#ffffff", "#b8b8b8", "#00e5ff", "#001114",
             "#7fb0ff", "#ff8080", "#d0a0ff", "#333333",
             "#ffe14d"),
    }

    BG = PANEL = PANEL2 = FIELD = FG = DIM = ACCENT = ACC_FG = None
    OWNEDC = MISSC = QUARC = SEL = ACCENT2 = None
    HOVER = SELHL = ACC2_FG = DANGER = DANGER_HOV = DANGER_FG = None
    THEME_NAME = "Boxcar Slate"

    def _mixc(a, b, t):
        ah = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
        bh = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
        return "#" + "".join(f"{round(x + (y - x) * t):02x}"
                             for x, y in zip(ah, bh))

    def _lumc(c):
        r, g, b = (int(c[i:i + 2], 16) / 255 for i in (1, 3, 5))
        return 0.2126 * r + 0.7152 * g + 0.0722 * b

    def apply_palette(name):
        global BG, PANEL, PANEL2, FIELD, FG, DIM, ACCENT, ACC_FG
        global OWNEDC, MISSC, QUARC, SEL, ACCENT2, STATUS_COLOR, THEME_NAME
        p = PALETTES.get(name) or PALETTES["Boxcar Slate"]
        THEME_NAME = name if name in PALETTES else "Boxcar Slate"
        (BG, PANEL, PANEL2, FIELD, FG, DIM, ACCENT, ACC_FG,
         OWNEDC, MISSC, QUARC, SEL, ACCENT2) = p
        global HOVER, SELHL, ACC2_FG, DANGER, DANGER_HOV, DANGER_FG
        HOVER = _mixc(PANEL2, ACCENT2, 0.30)
        SELHL = _mixc(BG, ACCENT2, 0.38)
        ACC2_FG = "#12141a" if _lumc(ACCENT2) > 0.45 else "#f4f6f9"
        DANGER = _mixc(BG, MISSC, 0.42)
        DANGER_HOV = MISSC
        DANGER_FG = "#12141a" if _lumc(MISSC) > 0.45 else "#f4f6f9"
        STATUS_COLOR = {"installed": ACCENT, "owned": OWNEDC,
                        "quarantined": QUARC, "missing": MISSC}

    apply_palette("Boxcar Slate")

    STATUS_LABEL = {"installed": "Installed",
                    "owned": "Owned (not installed)",
                    "quarantined": "Disabled",
                    "missing": "Not owned"}
    STATUS_ORDER = {"installed": 0, "owned": 1, "quarantined": 2, "missing": 3}

    CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


    def attach_tip(w, text):
        """Small hover tooltip; text may be a str or a callable."""
        st = {"win": None, "after": None}

        def _show():
            st["after"] = None
            try:
                t = text() if callable(text) else text
            except Exception:
                t = ""
            if not t:
                return
            tip = tk.Toplevel(w)
            tip.wm_overrideredirect(True)
            try:
                tip.attributes("-topmost", True)
            except Exception:
                pass
            tk.Label(tip, text=t, bg="#15181c", fg="#e8edf2",
                     font=("Segoe UI", 9), justify="left", wraplength=340,
                     padx=9, pady=6).pack()
            tip.wm_geometry(f"+{w.winfo_rootx() + 12}"
                            f"+{w.winfo_rooty() + w.winfo_height() + 6}")
            st["win"] = tip

        def _hide(_e=None):
            if st["after"]:
                try:
                    w.after_cancel(st["after"])
                except Exception:
                    pass
                st["after"] = None
            if st["win"]:
                try:
                    st["win"].destroy()
                except Exception:
                    pass
                st["win"] = None

        def _enter(_e=None):
            _hide()
            st["after"] = w.after(600, _show)

        w.bind("<Enter>", _enter, add="+")
        w.bind("<Leave>", _hide, add="+")
        w.bind("<ButtonPress>", _hide, add="+")


    class Gui(tk.Tk):
        def __init__(self, config_path=None):
            super().__init__()
            try:
                _f = float(load_json(DATA_DIR / "config.json", {}
                                     ).get("ui_scale", 1.0))
                if _f != 1.0:
                    self.tk.call("tk", "scaling",
                                 float(self.tk.call("tk", "scaling")) * _f)
                self._ui_scale = _f
            except Exception:
                self._ui_scale = 1.0
            self.config_path = config_path
            self.q = queue.Queue()
            self.state = {}
            self.quar = set()
            self.rows = {}            # pid -> row dict
            self.busy = False
            self.sort_col, self.sort_rev = "status", False

            self.title("Run8 DLC Manager -- DEMO (read-only)"
                       if DEMO_MODE else "Run8 DLC Manager")
            self.minsize(1000, 640)
            geo = load_json(DATA_DIR / "config.json", {}).get("geometry")
            applied = False
            if geo:
                try:
                    import re as _re
                    m = _re.fullmatch(r"(\d+)x(\d+)([+-]\d+)([+-]\d+)",
                                      geo)
                    if m:
                        w, h = int(m.group(1)), int(m.group(2))
                        x, y = int(m.group(3)), int(m.group(4))
                        try:
                            import ctypes
                            u = ctypes.windll.user32
                            vx, vy = (u.GetSystemMetrics(76),
                                      u.GetSystemMetrics(77))
                            vw, vh = (u.GetSystemMetrics(78),
                                      u.GetSystemMetrics(79))
                        except Exception:
                            vx, vy = 0, 0
                            vw = self.winfo_screenwidth()
                            vh = self.winfo_screenheight()
                        if (w >= 900 and h >= 600
                                and vx - 50 <= x < vx + vw - 200
                                and vy - 50 <= y < vy + vh - 200):
                            self.geometry(geo)
                            applied = True
                except Exception:
                    pass
            if not applied:
                self.geometry("1360x860")
            self.configure(bg=BG)
            try:
                try:
                    import base64 as _b64
                    import struct as _st
                    write_theme_icon(DATA_DIR / "run8dlc.ico")
                    self.iconbitmap(str(DATA_DIR / "run8dlc.ico"))
                    raw = _b64.b64decode(
                        THEME_ICONS.get(THEME_NAME)
                        or THEME_ICONS["Boxcar Slate"])
                    n = _st.unpack("<HHH", raw[:6])[2]
                    self._iconimgs = []
                    for i in range(n):
                        e = _st.unpack("<BBBBHHII",
                                       raw[6 + 16 * i:22 + 16 * i])
                        if (e[0] or 256) in (32, 64):
                            png = raw[e[7]:e[7] + e[6]]
                            self._iconimgs.append(tk.PhotoImage(
                                data=_b64.b64encode(png)))
                    if self._iconimgs:
                        self.iconphoto(True, *self._iconimgs)
                except Exception:
                    pass
            except Exception:
                pass

            self._fonts()
            self._style()
            self._build()
            self._proc = None
            self.protocol("WM_DELETE_WINDOW", self._on_close)
            self.after(80, self._poll)
            self.log(f"Run8 DLC Manager v{VERSION} -- catalog dated "
                     f"{self.app().catalog_date}")
            if DEMO_MODE:
                self.log("DEMO MODE -- statuses are fake and every "
                         "action is SIMULATED. Try everything; your "
                         "real library is never touched. Check My "
                         "Collection resets the demo.", "warn")
            if (DATA_DIR / "run8dlc_gui.py").exists():
                self.log("old-version files detected in this folder -- "
                         "delete run8dlc.py / run8dlc_gui.py and recreate "
                         "your shortcuts via Setup, or they may keep "
                         "launching the old app", "warn")
            has_imgs = bool(list((DATA_DIR / "media").glob("*_t.png")))
            self._set_mode("Gallery" if has_imgs else "List")
            if load_json(DATA_DIR / "config.json", {}).get("_setup_done"):
                self.rescan()
            else:
                self.after(150, lambda: SetupWizard(self, first_run=True))

        # ------------------------------------------------------------- helpers

        def app(self):
            """Fresh core App (re-reads ledger/config/catalog each time)."""
            return App(self.config_path)

        def log(self, line, color=None):
            self.q.put(("log", line, color))

        def _poll(self):
            try:
                while True:
                    item = self.q.get_nowait()
                    if item[0] == "log":
                        self._log_ui(item[1], item[2])
                    elif item[0] == "data":
                        self._render(item[1], item[2], item[3], item[4])
                    elif item[0] == "busy":
                        self._set_busy(item[1])
                    elif item[0] == "call":
                        item[1]()
                    elif item[0] == "progress":
                        if item[1] is None:
                            self.dlbar.pack_forget()
                        else:
                            if not self.dlbar.winfo_ismapped():
                                self.dlbar.pack(fill="x", pady=(4, 0))
                            self.dlbar["value"] = item[1]
            except queue.Empty:
                pass
            self.after(80, self._poll)

        def _set_busy(self, b):
            self.busy = b
            state = "disabled" if b else "normal"
            for w in self.toolbar_btns:
                w.configure(state=state)
            self.b_stop.configure(state="normal" if b else "disabled")
            if not b:
                self._on_select()

        # -------------------------------------------------------------- theme

        def _fonts(self):
            base = "Segoe UI" if sys.platform == "win32" else "TkDefaultFont"
            for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                         "TkHeadingFont"):
                try:
                    tkfont.nametofont(name).configure(
                        family="Segoe UI" if sys.platform == "win32" else None,
                        size=10)
                except Exception:
                    pass
            self.f_title = tkfont.Font(family=base, size=14, weight="bold")
            self.f_big = tkfont.Font(family=base, size=12, weight="bold")
            self.f_small = tkfont.Font(family=base, size=8)
            self.f_ui = tkfont.Font(family=base, size=10)
            self.f_mono = tkfont.Font(
                family="Consolas" if sys.platform == "win32" else "TkFixedFont",
                size=9)
            self.f_hdr = tkfont.Font(family=base, size=16, weight="bold")

        def _style(self):
            s = ttk.Style(self)
            s.theme_use("clam")
            s.configure(".", background=BG, foreground=FG, fieldbackground=FIELD,
                        bordercolor=PANEL2, lightcolor=PANEL, darkcolor=PANEL)
            s.configure("TFrame", background=BG)
            s.configure("Panel.TFrame", background=PANEL)
            s.configure("TLabel", background=BG, foreground=FG)
            s.configure("Panel.TLabel", background=PANEL, foreground=FG)
            s.configure("Dim.TLabel", background=PANEL, foreground=DIM)
            s.configure("Title.TLabel", background=BG, foreground=ACCENT)
            s.configure("TButton", background=PANEL2, foreground=FG,
                        borderwidth=0, focusthickness=0, padding=(10, 6))
            s.map("TButton",
                  background=[("active", HOVER), ("disabled", PANEL)],
                  foreground=[("disabled", "#5a636f")])
            s.configure("Accent.TButton", background=ACCENT,
                        foreground=ACC_FG)
            s.map("Accent.TButton",
                  background=[("active", ACCENT2), ("disabled", PANEL2)],
                  foreground=[("active", ACC2_FG),
                              ("disabled", "#5a636f")])
            s.configure("Danger.TButton", background=DANGER,
                        foreground=DANGER_FG)
            s.map("Danger.TButton",
                  background=[("active", DANGER_HOV), ("disabled", PANEL)],
                  foreground=[("active", DANGER_FG)])
            s.configure("TEntry", insertcolor=FG, padding=6)
            s.configure("TCombobox", padding=4, arrowcolor=FG)
            s.map("TCombobox", fieldbackground=[("readonly", FIELD)],
                  foreground=[("readonly", FG)])
            self.option_add("*TCombobox*Listbox.background", PANEL2)
            self.option_add("*TCombobox*Listbox.foreground", FG)
            self.option_add("*TCombobox*Listbox.selectBackground", ACCENT2)
            self.option_add("*TCombobox*Listbox.selectForeground", ACC2_FG)
            s.configure("Treeview", background=PANEL, fieldbackground=PANEL,
                        foreground=FG,
                        rowheight=int(27 * getattr(self, "_ui_scale",
                                                   1.0)),
                        borderwidth=0)
            s.configure("Treeview.Heading", background=PANEL2, foreground=DIM,
                        relief="flat", padding=(8, 6))
            s.map("Treeview.Heading", background=[("active", HOVER)])
            s.map("Treeview", background=[("selected", SELHL)],
                  foreground=[("selected", FG)])
            s.configure("Vertical.TScrollbar", background=PANEL2,
                        troughcolor=BG, borderwidth=0, arrowcolor=DIM)

        # -------------------------------------------------------------- build

        def _build(self):
            # toolbar ---------------------------------------------------------
            bar = ttk.Frame(self, padding=(14, 12, 14, 8))
            bar.pack(fill="x")
            self.masthead = tk.Canvas(bar, highlightthickness=0, bg=BG)
            self.masthead.pack(side="left")
            try:
                draw_title_train(self.masthead, THEME_NAME, self.f_title,
                                 getattr(self, "_ui_scale", 1.0))
            except Exception as e:
                self.masthead.create_text(6, 20, anchor="w",
                                          text="RUN8 DLC MANAGER",
                                          fill=ACCENT, font=self.f_title)
            self.v_view = tk.StringVar(value="Gallery")
            self.b_tab_c = tk.Button(bar, text="  Gallery  ", bd=0,
                                     relief="flat", cursor="hand2",
                                     bg=ACCENT, fg=ACC_FG,
                                     activebackground=ACCENT,
                                     font=self.f_big,
                                     command=lambda: self._set_mode("Gallery"))
            self.b_tab_m = tk.Button(bar, text="  List  ", bd=0,
                                     relief="flat", cursor="hand2",
                                     bg=PANEL2, fg=DIM,
                                     activebackground=PANEL2,
                                     font=self.f_big,
                                     command=lambda: self._set_mode("List"))
            self.b_tab_s = tk.Button(bar, text="  Settings  ", bd=0,
                                     relief="flat", cursor="hand2",
                                     bg=PANEL2, fg=DIM,
                                     activebackground=PANEL2,
                                     font=self.f_big,
                                     command=lambda:
                                     self._set_mode("Settings"))
            self.b_tab_c.pack(side="left", padx=(16, 2), pady=2)
            self.b_tab_m.pack(side="left", padx=(0, 2), pady=2)
            self.b_tab_s.pack(side="left", padx=(0, 8), pady=2)
            self.toolbar_btns = []

            def tb(text, cmd, style="TButton", pad=6):
                b = ttk.Button(bar, text=text, command=cmd, style=style)
                b.pack(side="right", padx=(pad, 0))
                self.toolbar_btns.append(b)
                return b


            tb("Check My Collection", self.rescan)
            tb("Add Purchase…", self.add_purchase, "Accent.TButton")

            # filter row ------------------------------------------------------
            fr = ttk.Frame(self, padding=(14, 0, 14, 8))
            fr.pack(fill="x")
            ttk.Label(fr, text="Search", foreground=ACCENT,
                      font=self.f_big).pack(side="left")
            self.v_search = tk.StringVar()
            self.v_search.trace_add("write", lambda *_: self._refilter())
            e = ttk.Entry(fr, textvariable=self.v_search, width=40)
            e.pack(side="left", padx=(8, 18))
            ttk.Style(self).configure("Ph.TEntry", foreground=DIM)
            self._ph_active = False
            _PH = "Type a route or engine name…"

            def _ph_show(_e=None):
                if not self.v_search.get():
                    self._ph_active = True
                    e.configure(style="Ph.TEntry")
                    self.v_search.set(_PH)

            def _ph_hide(_e=None):
                if self._ph_active:
                    self._ph_active = False
                    self.v_search.set("")
                e.configure(style="TEntry")

            e.bind("<FocusIn>", _ph_hide, add="+")
            e.bind("<FocusOut>", _ph_show, add="+")
            self.after(50, _ph_show)
            self.v_status = tk.StringVar(value="All")
            ttk.Label(fr, text="Category", foreground=ACCENT,
                      font=self.f_big).pack(side="left")
            self.v_cat = tk.StringVar(value="All")
            cb2 = ttk.Combobox(fr, textvariable=self.v_cat, state="readonly",
                               width=14, values=["All", "Route", "Locomotive",
                                                 "Passenger", "Rolling Stock"])
            cb2.pack(side="left", padx=(8, 0))
            cb2.bind("<<ComboboxSelected>>", lambda *_: self._refilter())
            self.lbl_summary = ttk.Label(fr, text="", style="Dim.TLabel")
            self.lbl_summary.pack(side="right")

            # one-click status chips -------------------------------------
            ch = ttk.Frame(self, padding=(14, 0, 14, 8))
            ch.pack(fill="x")
            self._chips = {}

            def _chip(lbl, val):
                b = tk.Button(ch, text=f"  {lbl}  ", bd=0, relief="flat",
                              cursor="hand2", font=self.f_big, bg=PANEL2,
                              fg=DIM, activebackground=ACCENT,
                              activeforeground=ACC_FG,
                              command=lambda: (self.v_status.set(val),
                                               self._refilter()))
                b.pack(side="left", padx=(0, 8))
                self._chips[val] = b

            ttk.Label(ch, text="Status", foreground=ACCENT,
                      font=self.f_big).pack(side="left", padx=(0, 10))
            for _lbl, _val in (("All", "All"), ("Installed", "Installed"),
                               ("Owned (not installed)",
                                "Owned (not installed)"),
                               ("Disabled", "Disabled"),
                               ("Not owned", "Not owned")):
                _chip(_lbl, _val)

            # main area -------------------------------------------------------
            main = ttk.Frame(self, padding=(14, 0, 14, 0))
            main.pack(fill="both", expand=True)
            main.columnconfigure(0, weight=1)
            main.rowconfigure(0, weight=1)

            # List view lives in a paned window -- drag the divider
            # between the table and the info panel to taste
            self.pw = ttk.PanedWindow(main, orient="horizontal")
            lf = ttk.Frame(self.pw)
            lf.columnconfigure(0, weight=1)
            lf.rowconfigure(0, weight=1)
            cols = ("name", "category", "status", "price", "evidence")
            self.tree = ttk.Treeview(lf, columns=cols, show="headings",
                                     selectmode="browse")
            heads = {"name": ("Product", 330), "category": ("Category", 110),
                     "status": ("Status", 150), "price": ("Price", 70),
                     "evidence": ("How I Know", 210)}
            for c in cols:
                self.tree.heading(c, text=heads[c][0],
                                  command=lambda cc=c: self._sort_by(cc))
                self.tree.column(c, width=heads[c][1],
                                 anchor="w" if c != "price" else "e",
                                 stretch=(c in ("name", "evidence")))
            for st, col in STATUS_COLOR.items():
                self.tree.tag_configure(st, foreground=col)
            self.tree.grid(row=0, column=0, sticky="nsew")
            self.tree_sb = sb = ttk.Scrollbar(lf, orient="vertical",
                                              command=self.tree.yview)
            self.tree.configure(yscrollcommand=sb.set)
            sb.grid(row=0, column=1, sticky="ns")
            self.tree.bind("<<TreeviewSelect>>", lambda *_: self._on_select())
            self.tree.bind("<Double-1>", lambda *_: self.open_store())

            self.gal = tk.Canvas(main, bg=PANEL, highlightthickness=0,
                                 yscrollincrement=1)
            self._gal_sb = ttk.Scrollbar(main, orient="vertical",
                                         command=self.gal.yview)
            self.gal.configure(yscrollcommand=self._gal_sb.set)
            self.gal.bind("<Configure>", self._gal_resized)
            self.gal.bind_all("<MouseWheel>", self._gal_scroll)
            self._gtiles, self._gorder = {}, []

            # detail panel ----------------------------------------------------
            self.det = det = ttk.Frame(self.pw, style="Panel.TFrame",
                                       padding=14, width=620)
            self.pw.add(lf, weight=1)
            self.pw.add(det, weight=0)
            det.grid_propagate(False)
            det.bind("<Configure>", self._det_resized, add="+")
            self.d_name = ttk.Label(det, text="Select a product",
                                    style="Panel.TLabel", font=self.f_hdr,
                                    foreground=ACCENT,
                                    wraplength=390, justify="left")
            self.d_name.pack(anchor="w")
            self.d_meta = ttk.Label(det, text="", style="Dim.TLabel")
            self.d_meta.pack(anchor="w", pady=(2, 0))
            self.d_status = tk.Label(det, text="", bg=PANEL, fg=DIM,
                                     font=self.f_big, anchor="w")
            self.d_status.pack(anchor="w", pady=(8, 4), fill="x")
            info = ttk.Frame(det, style="Panel.TFrame")
            info.pack(fill="x", pady=(0, 8))
            info.columnconfigure(0, weight=1)
            self.d_desc = tk.Label(info, text="", bg=PANEL, fg="#a8b2c0",
                                   wraplength=205, justify="left",
                                   anchor="nw")
            self.d_desc.grid(row=0, column=0, sticky="nw")
            self.d_ev = tk.Text(info, height=9, width=26, bg=FIELD, fg=DIM,
                                bd=0, font=self.f_mono, wrap="word",
                                highlightthickness=0, padx=6, pady=6)
            self.d_ev.grid(row=0, column=1, sticky="ne", padx=(10, 0))
            self.d_ev.configure(state="disabled")

            self.d_img = tk.Label(det, bg=FIELD, fg=DIM, anchor="center",
                                  text="")
            self.d_img.pack(fill="both", expand=True, pady=(2, 0))
            self._imgs, self._thumbs = {}, {}

            bt = ttk.Frame(det, style="Panel.TFrame")
            bt.pack(fill="x", pady=(10, 0))
            bt.columnconfigure(0, weight=1)
            bt.columnconfigure(1, weight=1)
            self.b_install = ttk.Button(bt, text="Reinstall",
                                        command=self.reinstall)
            self.b_install.grid(row=0, column=0, sticky="we",
                                padx=(0, 3), pady=2)
            self.b_uninst = ttk.Button(bt, text="Disable (keeps the files)",
                                       style="Danger.TButton",
                                       command=self.uninstall)
            self.b_uninst.grid(row=0, column=1, sticky="we",
                               padx=(3, 0), pady=2)
            self.b_restore = ttk.Button(bt, text="Enable (put back)",
                                        command=self.restore)
            self.b_restore.grid(row=1, column=0, sticky="we",
                                padx=(0, 3), pady=2)
            self.b_store = ttk.Button(bt, text="Open store page",
                                      command=self.open_store)
            self.b_store.grid(row=1, column=1, sticky="we",
                              padx=(3, 0), pady=2)
            self.b_buy = ttk.Button(bt, text="Buy on 3DTS store",
                                    style="Accent.TButton",
                                    command=self.open_store)
            self.b_buy.grid(row=2, column=0, columnspan=2, sticky="we",
                            pady=2)

            # log + status bar ------------------------------------------------
            lg = ttk.Frame(self, padding=(14, 10, 14, 4))
            lg.pack(fill="x")
            self.logbox = tk.Text(lg, height=7, bg=FIELD, fg=DIM, bd=0,
                                  font=self.f_mono, wrap="word",
                                  highlightthickness=0, padx=8, pady=6)
            self.logbox.pack(fill="x")
            self.logbox.tag_configure("accent", foreground=ACCENT)
            self.logbox.tag_configure("warn", foreground=MISSC)
            self.logbox.configure(state="disabled")
            s = ttk.Style(self)
            s.configure("Dl.Horizontal.TProgressbar", troughcolor=FIELD,
                        background=ACCENT, borderwidth=0)
            self.dlbar = ttk.Progressbar(lg, style="Dl.Horizontal.TProgressbar",
                                         maximum=100)
            # packed on demand while a download is running
            sbrow = ttk.Frame(self)
            sbrow.pack(fill="x")
            self.statusbar = ttk.Label(sbrow, text="",
                                       padding=(16, 4, 16, 8),
                                       foreground=DIM)
            self.statusbar.pack(side="left", fill="x", expand=True)
            self.b_stop = ttk.Button(sbrow, text="Stop",
                                     style="Danger.TButton",
                                     command=self._stop_op,
                                     state="disabled")
            self.b_stop.pack(side="right", padx=(0, 14), pady=(0, 6))

            # hover tips ------------------------------------------------
            attach_tip(self.b_tab_c,
                       "Every product with its picture, like a catalog")
            attach_tip(self.b_tab_m,
                       "The same products as a sortable table")
            attach_tip(self.b_tab_s,
                       "Folders, appearance, tools, import and backup "
                       "-- all in one page")
            _tips = {"Check My Collection":
                     "Scan your folders and the game to update what's "
                     "installed and owned",
                     "Add Purchase…":
                     "Record a purchase -- automatic from the receipt "
                     "page, or fully manual with a transaction ID or "
                     "installer file"}
            for _b in self.toolbar_btns:
                _t = _tips.get(str(_b.cget("text")))
                if _t:
                    attach_tip(_b, _t)
            attach_tip(e, "Type part of a name -- e.g. 'cajon' or 'sd40'")
            attach_tip(cb2, "Show only routes, engines, or cars")
            attach_tip(self.b_install, lambda:
                       "Launch the saved installer for this product"
                       if str(self.b_install.cget("state")) != "disabled"
                       else "No installer found on disk -- check the "
                            "Installers folder in Setup")
            def _uninst_tip():
                if str(self.b_uninst.cget("state")) != "disabled":
                    return ("Move this route's folders to a safe "
                            "place -- nothing is deleted")
                r = self._sel()
                if r and r["status"] == "installed":
                    return ("Equipment packs share their files with "
                            "other products, so only ROUTES can be "
                            "disabled individually")
                return "Only available when the product is installed"
            attach_tip(self.b_uninst, _uninst_tip)
            attach_tip(self.b_restore, lambda:
                       "Put the disabled product straight back into "
                       "the game"
                       if str(self.b_restore.cget("state")) != "disabled"
                       else "Nothing is disabled for this product")
            attach_tip(self.b_buy,
                       "Opens the product's page on the real 3DTS store "
                       "in your browser")
            attach_tip(self.b_store,
                       "Opens the product's page on the 3DTS store "
                       "in your browser")
            attach_tip(self.b_stop, "Stop the current operation")

        # ------------------------------------------------------------ log ui

        def _log_ui(self, line, color=None):
            self.logbox.configure(state="normal")
            self.logbox.insert("end", line.rstrip() + "\n",
                               (color,) if color else ())
            self.logbox.see("end")
            self.logbox.configure(state="disabled")

        # ------------------------------------------------------------- rescan

        def _stop_op(self):

            p = getattr(self, "_proc", None)

            if p and p.poll() is None:

                p.terminate()

                self.log("stopped by user", "warn")


        def _on_close(self):
            try:
                cfg = load_json(DATA_DIR / "config.json", {})
                cfg["geometry"] = self.geometry()
                try:
                    if self.pw.winfo_ismapped():
                        w_ = max(self.pw.winfo_width(), 1)
                        cfg["det_split"] = round(
                            min(0.9, max(0.1,
                                         self.pw.sashpos(0) / w_)), 3)
                    elif getattr(self, "_split", None):
                        cfg["det_split"] = round(self._split, 3)
                    cfg.pop("det_sash", None)
                except Exception:
                    pass
                save_json(DATA_DIR / "config.json", cfg)
            except Exception:
                pass
            self._stop_op()
            self.destroy()


        def rescan(self):
            if self.busy:
                return
            self.q.put(("busy", True))
            self.log("scanning library…")

            def work():
                try:
                    app = self.app()
                    state, unmatched = app.build_state(include_game_scan=True)
                    quar = quarantined_pids()
                    rows = {}
                    for p in app.products:
                        if p.get("category") == "Base":
                            continue
                        st = product_status(p["id"], state, quar)
                        ev = state.get(p["id"],
                                       {"ledger": [], "installers": [],
                                        "receipts": [], "game": [],
                                        "game_files": []})
                        bits = []
                        if ev["installers"]:
                            bits.append(f"{len(ev['installers'])} exe")
                        if ev["receipts"]:
                            bits.append("receipt")
                        if ev["game"]:
                            bits.append("in-game")
                        elif ev.get("game_files"):
                            bits.append("in-game files")
                        if ev["ledger"]:
                            bits.append("ledger")
                        rows[p["id"]] = {
                            "pid": p["id"], "name": p["name"],
                            "category": p.get("category", ""),
                            "status": st, "price": p.get("price"),
                            "url": p.get("url"), "ev": ev,
                            "evtext": " · ".join(bits) if bits else "—",
                        }
                    if DEMO_MODE:
                        import zlib
                        _mix = ("installed", "installed", "installed",
                                "installed", "missing", "missing",
                                "missing", "missing", "owned",
                                "quarantined")
                        for _r in rows.values():
                            _r["status"] = _mix[zlib.crc32(
                                _r["pid"].encode()) % len(_mix)]
                            _r["evtext"] = "demo data"
                            _r["ev"] = {"ledger": [], "installers": [],
                                        "receipts": [], "game": [],
                                        "game_files": []}
                        quar = {_p for _p, _r in rows.items()
                                if _r["status"] == "quarantined"}
                    self.q.put(("data", rows, state, quar, unmatched))
                except Exception as e:
                    self.log(f"rescan failed: {e!r}", "warn")
                finally:
                    self.q.put(("busy", False))

            threading.Thread(target=work, daemon=True).start()

        def _render(self, rows, state, quar, unmatched):
            self.rows, self.state, self.quar = rows, state, quar
            self._refilter()
            n = {"installed": 0, "owned": 0, "quarantined": 0, "missing": 0}
            cost, unknown = 0.0, 0
            for r in rows.values():
                n[r["status"]] += 1
                if r["status"] == "missing":
                    if r["price"] is None:
                        unknown += 1
                    else:
                        cost += r["price"]
            c = f"${cost:,.2f}" + (f" + {unknown} unpriced" if unknown else "")
            self.statusbar.configure(
                text=f"{len(rows)} DLC products   ·   {n['installed']} installed"
                     f"   ·   {n['owned']} owned (not installed)"
                     f"   ·   {n['quarantined']} disabled"
                     f"   ·   {n['missing']} not owned   ·   "
                     f"complete the set: {c}")
            self.log(f"scan done: {n['installed']} installed, {n['owned']} owned, "
                     f"{n['missing']} not owned", "accent")
            try:
                inst_now = {p_ for p_, r_ in rows.items()
                            if r_["status"] == "installed"}
                prev = getattr(self, "_prev_installed", None)
                self._prev_installed = inst_now
                fresh_pids = (inst_now - prev) if prev is not None else set()
                if fresh_pids and not DEMO_MODE:
                    names = ", ".join(sorted(
                        rows[p_]["name"] for p_ in fresh_pids)[:4])
                    if len(fresh_pids) > 4:
                        names += ", …"
                    if messagebox.askyesno(
                            "New DLC installed",
                            f"Now installed: {names}.\n\nRun the "
                            "official Run8 updater? It updates DLC "
                            "files too, and is usually needed for "
                            "compatibility with the latest game "
                            "version."):
                        self._run_updater_stream()
            except Exception:
                pass
            self._check_game_update()
            if unmatched:
                self.log(f"{len(unmatched)} file(s) need review -- run 'report "
                         "--verbose' or add overrides to mapping.json", "warn")

        def _refilter(self):
            self.tree.delete(*self.tree.get_children())
            q = ("" if getattr(self, "_ph_active", False)
                 else squash(self.v_search.get()))
            wanted_status = self.v_status.get()
            wanted_cat = self.v_cat.get()
            for _v, _b in getattr(self, "_chips", {}).items():
                _on = _v == wanted_status
                _b.configure(bg=ACCENT if _on else PANEL2,
                             fg=ACC_FG if _on else DIM)
            items = list(self.rows.values())
            key = {"name": lambda r: r["name"].lower(),
                   "category": lambda r: (r["category"], r["name"].lower()),
                   "status": lambda r: (STATUS_ORDER[r["status"]],
                                        r["category"], r["name"].lower()),
                   "price": lambda r: (r["price"] is None, r["price"] or 0),
                   "evidence": lambda r: r["evtext"]}[self.sort_col]
            items.sort(key=key, reverse=self.sort_rev)
            shown = []
            for r in items:
                if q and q not in squash(r["name"]):
                    continue
                if wanted_status != "All" and \
                        STATUS_LABEL[r["status"]] != wanted_status:
                    continue
                if wanted_cat != "All" and r["category"] != wanted_cat:
                    continue
                self.tree.insert("", "end", iid=r["pid"], tags=(r["status"],),
                                 values=(r["name"], r["category"],
                                         STATUS_LABEL[r["status"]],
                                         price_str(r), r["evtext"]))
                shown.append(r)
            try:
                inst = sum(1 for r in shown if r["status"] == "installed")
                own = sum(1 for r in shown if r["status"] == "owned")
                miss = [r for r in shown if r["status"] == "missing"]
                cost = sum(r["price"] for r in miss
                           if r.get("price") is not None)
                s = f"{inst} installed · {own} owned · {len(miss)} missing"
                if miss and cost:
                    s += f" · complete for ${cost:,.0f}"
                self.lbl_summary.configure(text=s)
                if self.rows and not getattr(self, "_greeted", False):
                    self._greeted = True
                    tot = len(self.rows)
                    have = sum(1 for rr in self.rows.values()
                               if rr["status"] != "missing")
                    gm = f"Ready -- you own {have} of {tot} add-ons"
                    _mi = [rr for rr in self.rows.values()
                           if rr["status"] == "missing"]
                    _gc = sum(rr["price"] for rr in _mi
                              if rr.get("price") is not None)
                    if _mi:
                        gm += (f" -- the remaining {len(_mi)} would cost "
                               f"${_gc:,.0f}" if _gc else
                               f" -- {len(_mi)} more on the store")
                    self.log(gm, "accent")
            except Exception:
                pass
            if shown and not self.tree.selection():
                try:
                    self.tree.selection_set(shown[0]["pid"])
                    self.tree.focus(shown[0]["pid"])
                except Exception:
                    pass
            if self.v_view.get() == "Gallery":
                key = (tuple((r["pid"], r["status"]) for r in shown),
                       self._gal_metrics()[:2])
                if key != getattr(self, "_gal_key", None):
                    self._gal_key = key
                    self._build_gallery(shown)
            else:
                self._autosize_columns(shown)

        def _set_mode(self, mode):
            if mode in ("Gallery", "List"):
                self._last_view = mode
            self.v_view.set(mode)
            for b, m in ((self.b_tab_c, "Gallery"),
                         (self.b_tab_m, "List"),
                         (self.b_tab_s, "Settings")):
                on = mode == m
                b.configure(bg=ACCENT if on else PANEL2,
                            fg=ACC_FG if on else DIM)
            self._swap_view()

        def _det_resized(self, ev):
            # labels re-wrap to whatever width the panel actually has
            try:
                w = max(ev.width, 320)
                self.d_name.configure(wraplength=w - 40)
                self.d_desc.configure(wraplength=max(205, w - 380))
                r = getattr(self, "_det_r", None)
                if r:
                    self._load_page(r)
            except Exception:
                pass

        def _set_sash(self, sp):
            try:
                self.pw.sashpos(0, int(sp))
            except Exception:
                pass

        def _swap_view(self):
            try:
                self._close_card()
                sp_ = getattr(self, "_setpane", None)
                if sp_ is not None:
                    sp_.destroy()
                    self._setpane = None
                try:
                    if self.pw.winfo_ismapped():
                        w_ = max(self.pw.winfo_width(), 1)
                        self._split = min(0.9, max(
                            0.1, self.pw.sashpos(0) / w_))
                except Exception:
                    pass
                mode = self.v_view.get()
                if mode in ("Settings", "Add", "Import",
                            "Transactions"):
                    self.pw.grid_remove()
                    self.gal.grid_remove()
                    self._gal_sb.grid_remove()
                    if mode == "Settings":
                        self._setpane = SetupPane(
                            self.gal.master, self, first_run=False,
                            on_done=lambda: (self.log("settings saved",
                                                      "accent"),
                                             self.rescan()))
                    else:
                        f = tk.Frame(self.gal.master, bg=PANEL,
                                     padx=20, pady=16)
                        {"Add": self._pane_add_purchase,
                         "Import": self._pane_import,
                         "Transactions": self._pane_transactions
                         }[mode](f)
                        self._setpane = f
                    self._setpane.grid(row=0, column=0, columnspan=3,
                                       sticky="nsew")
                elif mode == "Gallery":
                    self.pw.grid_remove()
                    self.gal.grid(row=0, column=0, sticky="nsew")
                    self._gal_sb.grid(row=0, column=1, sticky="ns")
                else:
                    self.gal.grid_remove()
                    self._gal_sb.grid_remove()
                    self.pw.grid(row=0, column=0, columnspan=2,
                                 sticky="nsew")
                    if not hasattr(self, "_split"):
                        ratio = load_json(DATA_DIR / "config.json",
                                          {}).get("det_split", 0.6)
                        try:
                            ratio = float(ratio)
                        except Exception:
                            ratio = 0.6
                        self._split = (ratio if 0.1 <= ratio <= 0.9
                                       else 0.6)
                    self.after(80, lambda: self._set_sash(
                        max(320, int(self.pw.winfo_width()
                                     * self._split))))
                self.update_idletasks()
                if mode in ("Gallery", "List"):
                    self._refilter()
            except Exception as e:
                self.log(f"view switch failed: {e!r}", "warn")

        def _gal_resized(self, *_):
            if getattr(self, "_gal_after", None):
                self.after_cancel(self._gal_after)
            self._gal_after = self.after(140, self._gal_reflow)

        def _gal_reflow(self):
            self._gal_after = None
            try:
                if self.v_view.get() != "Gallery":
                    return
                self._gal_key = None
                self._refilter()
            except Exception as e:
                self.log(f"gallery reflow failed: {e!r}", "warn")

        def _gal_scroll(self, ev):
            if not (self.v_view.get() == "Gallery"
                    and self.gal.winfo_ismapped()):
                return
            # accumulate the wheel into a target and glide there
            self._gal_target = (getattr(self, "_gal_target", 0)
                                - (ev.delta / 120) * 170)
            if not getattr(self, "_gal_gliding", False):
                self._gal_gliding = True
                self._gal_glide()

        def _gal_glide(self):
            t = getattr(self, "_gal_target", 0)
            step = int(t * 0.38)
            if step == 0:
                rem = int(round(t))
                if rem:
                    self.gal.yview_scroll(rem, "units")
                self._gal_target = 0
                self._gal_gliding = False
                return
            self._gal_target = t - step
            self.gal.yview_scroll(step, "units")
            self.after(15, self._gal_glide)

        def _build_gallery(self, items):
          try:
            c = self.gal
            c.delete("all")
            self._gtiles, self._gorder = {}, []
            _cols, tile_w, _th, ab = self._gal_metrics()
            if ab != getattr(self, "_last_ab", None):
                self._thumbs = {k: v for k, v in self._thumbs.items()
                                if isinstance(k, tuple) and k[1] == ab}
            self._last_ab = ab
            name_size = max(9, min(14, 9 + (tile_w - 200) // 55))
            self.f_tile = tkfont.Font(font=self.f_ui)
            self.f_tile.configure(size=name_size)
            self.f_price = tkfont.Font(font=self.f_ui)
            self.f_price.configure(size=name_size + 3, weight="bold")
            self.f_chip = tkfont.Font(font=self.f_ui)
            self.f_chip.configure(size=max(8, name_size - 1))
            wrapw = tile_w - 36
            from collections import Counter
            _sizes = Counter()
            for r in items:
                _t = self._thumb_scaled(r["pid"], ab)
                if _t:
                    _sizes[(_t.width(), _t.height())] += 1
            if _sizes:
                box_w, box_h = _sizes.most_common(1)[0][0]
            else:
                box_w = int(self.HERO_W * ab[0] / ab[1])
                box_h = int(box_w * 0.52)
            inner_w = tile_w - 12
            for r in items:
                pid, st = r["pid"], r["status"]
                tag = "t_" + pid
                sel = tag + "_sel"
                d = {"pid": pid, "tag": tag, "x": 0, "y": 0,
                     "desc_id": None, "arrow": None}
                d["card"] = c.create_rectangle(
                    0, 0, inner_w, 10, fill=PANEL2,
                    outline=STATUS_COLOR[st], width=2,
                    tags=(tag, sel))
                yy = 10
                cx = inner_w // 2
                c.create_rectangle(cx - box_w // 2, yy,
                                   cx + box_w // 2, yy + box_h,
                                   fill="#000000", outline="",
                                   tags=(tag, sel))
                th = self._thumb_scaled(pid, ab)
                if th:
                    th = self._crop_box(pid, ab, th, box_w, box_h)
                    c.create_image(cx, yy + box_h // 2, image=th,
                                   tags=(tag, sel))
                else:
                    c.create_text(cx, yy + box_h // 2, text="no image",
                                  fill=DIM, font=self.f_chip,
                                  tags=(tag, sel))
                yy += box_h + 8
                dsc = (self.app().by_id.get(pid, {}) or {}
                       ).get("desc") or ""
                d["dsc"] = dsc
                t_id = c.create_text(12, yy, anchor="nw",
                                     text=r["name"], fill=FG,
                                     font=self.f_tile,
                                     width=wrapw - 20,
                                     tags=(tag, tag + "_ttl"))
                bb = c.bbox(t_id)
                if dsc:
                    d["arrow"] = c.create_text(
                        bb[2] + 7, bb[1], anchor="nw", text="▾",
                        fill=ACCENT2, font=self.f_tile,
                        tags=(tag, tag + "_tog"))
                yy = bb[3] + 6
                d["desc_y"] = yy
                line_h = self.f_chip.metrics("linespace")
                if st == "missing":
                    ptxt = (f"${r['price']:.0f}"
                            if r.get("price") is not None else "$?")
                    c.create_text(12, yy, anchor="nw", text=ptxt,
                                  fill=ACCENT2, font=self.f_price,
                                  tags=(tag, sel))
                    yy += self.f_price.metrics("linespace") + 8
                else:
                    c.create_text(12, yy, anchor="nw",
                                  text=STATUS_LABEL[st].upper(),
                                  fill=STATUS_COLOR[st],
                                  font=self.f_chip, tags=(tag, sel))
                    if r.get("price") is not None:
                        c.create_text(inner_w - 12, yy, anchor="ne",
                                      text=f"${r['price']:.0f}",
                                      fill=DIM, font=self.f_chip,
                                      tags=(tag, sel))
                    yy += line_h + 8
                bx = [12]

                def _btn(label, cb, yy=yy, bx=bx, tag=tag,
                         accent=False, danger=False):
                    w_ = self.f_chip.measure(label) + 24
                    btag = f"{tag}_b{bx[0]}"
                    c.create_rectangle(
                        bx[0], yy, bx[0] + w_, yy + 26,
                        fill=(ACCENT if accent else FIELD),
                        outline=(MISSC if danger else ""),
                        tags=(tag, btag))
                    c.create_text(bx[0] + w_ // 2, yy + 13,
                                  text=label,
                                  fill=(ACC_FG if accent else
                                        MISSC if danger else FG),
                                  font=self.f_chip, tags=(tag, btag))
                    c.tag_bind(btag, "<Button-1>",
                               lambda e, f_=cb: f_())
                    c.tag_bind(btag, "<Enter>", lambda e:
                               c.configure(cursor="hand2"))
                    c.tag_bind(btag, "<Leave>", lambda e:
                               c.configure(cursor=""))
                    bx[0] += w_ + 8

                nbtn = 0
                if st == "missing" and r.get("url"):
                    _btn("Buy on 3DTS",
                         lambda pid=pid: self._tile_buy(pid),
                         accent=True)
                    nbtn += 1
                if st in ("installed", "owned"):
                    _btn("Reinstall" if st == "installed"
                         else "Run installer",
                         lambda pid=pid:
                         self._card_act(pid, self.reinstall))
                    nbtn += 1
                if st == "installed" and (
                        DEMO_MODE or
                        bool(_game_paths_for(self.app(), pid,
                                             self.state)[0])):
                    _btn("Disable",
                         lambda pid=pid:
                         self._card_act(pid, self.uninstall),
                         danger=True)
                    nbtn += 1
                if st == "quarantined":
                    _btn("Enable",
                         lambda pid=pid:
                         self._card_act(pid, self.restore))
                    nbtn += 1
                yy += (26 + 10) if nbtn else 2
                d["h"] = yy + 4
                c.coords(d["card"], 0, 0, inner_w, d["h"])
                c.tag_bind(sel, "<Button-1>",
                           lambda e, pid=pid: self._gal_pick(pid))
                c.tag_bind(tag + "_ttl", "<Button-1>",
                           lambda e, pid=pid:
                           (self._gal_pick(pid),
                            self._toggle_desc(pid)))
                c.tag_bind(tag + "_tog", "<Button-1>",
                           lambda e, pid=pid:
                           self._toggle_desc(pid))
                for tg_ in (tag + "_ttl", tag + "_tog"):
                    c.tag_bind(tg_, "<Enter>", lambda e:
                               c.configure(cursor="hand2"))
                    c.tag_bind(tg_, "<Leave>", lambda e:
                               c.configure(cursor=""))
                self._gtiles[pid] = d
                self._gorder.append(pid)
            self._gal_relayout()
          except Exception as e:
            self.log(f"gallery failed: {e!r}", "warn")


        HERO_W = 560           # native {pid}.png width; tiles downscale
        SCALE_STEPS = ((1, 2), (3, 5), (2, 3), (3, 4), (5, 6), (1, 1),
                       (7, 6), (4, 3), (3, 2))

        def _gal_metrics(self):
            w = max(self.gal.winfo_width(), 200)
            cols = int(getattr(self, "_gal_cols", 0) or
                       load_json(DATA_DIR / "config.json", {}
                                 ).get("gal_cols", 3))
            self._gal_cols = cols
            tile_w = w // cols
            if tile_w < 195:
                cols = max(1, w // 195)
                tile_w = w // cols
            inner = tile_w - 28
            best = (1, 2)
            for a, b in self.SCALE_STEPS:
                if self.HERO_W * a / b <= inner:
                    best = (a, b)
            img_w = int(self.HERO_W * best[0] / best[1])
            tile_h = int(img_w * 0.52) + 62
            return cols, tile_w, tile_h, best

        def _thumb_scaled(self, pid, ab):
            key = (pid, ab)
            t = self._thumbs.get(key)
            if t is None:
                base_img = self._hero(pid)
                if not base_img:
                    # no hero yet -- fall back to the small thumb
                    f = DATA_DIR / "media" / (pid + "_t.png")
                    try:
                        t = tk.PhotoImage(file=str(f)) if f.exists() \
                            else False
                    except Exception:
                        t = False
                elif ab == (1, 1):
                    t = base_img
                else:
                    try:
                        a, b = ab
                        t = base_img.zoom(a, a) if a > 1 else base_img
                        if b > 1:
                            t = t.subsample(b, b)
                    except Exception:
                        t = base_img
                self._thumbs[key] = t
            return t

        def _crop_box(self, pid, ab, img, bw, bh):
            """Canvas items don't clip like widget frames did -- crop a
            too-big photo to the letterbox, center-weighted, cached."""
            w, h = img.width(), img.height()
            if w <= bw and h <= bh:
                return img
            key = (pid, ab, bw, bh)
            t = self._thumbs.get(key)
            if t is None:
                try:
                    t = tk.PhotoImage()
                    x0 = max(0, (w - bw) // 2)
                    y0 = max(0, (h - bh) // 2)
                    t.tk.call(t, "copy", img, "-from", x0, y0,
                              min(x0 + bw, w), min(y0 + bh, h))
                except Exception:
                    t = img
                self._thumbs[key] = t
            return t

        def _gal_relayout(self):
            """Position tiles row by row; existing items just MOVE."""
            if not getattr(self, "_gorder", None):
                self.gal.configure(scrollregion=(0, 0, 0, 0))
                return
            cols, tile_w, _th, _ab = self._gal_metrics()
            w = max(self.gal.winfo_width(), 200)
            mx = max(8, (w - cols * tile_w) // 2 + 4)
            y, i = 8, 0
            order = self._gorder
            while i < len(order):
                row = order[i:i + cols]
                hts = [self._gtiles[p_]["h"] for p_ in row]
                for j, p_ in enumerate(row):
                    d = self._gtiles[p_]
                    nx, ny = j * tile_w + mx, y
                    if nx != d["x"] or ny != d["y"]:
                        self.gal.move(d["tag"], nx - d["x"],
                                      ny - d["y"])
                        d["x"], d["y"] = nx, ny
                y += (max(hts) if hts else 0) + 14
                i += cols
            self.gal.configure(scrollregion=(0, 0, w, y + 6))

        def _toggle_desc(self, pid):
            d = (getattr(self, "_gtiles", {}) or {}).get(pid)
            if not d or not d.get("dsc"):
                return
            c = self.gal
            _cols, tile_w, _th, _ab = self._gal_metrics()
            if d["desc_id"]:
                bb = c.bbox(d["desc_id"])
                dh = -((bb[3] - bb[1]) + 8)
                c.delete(d["desc_id"])
                d["desc_id"] = None
                if d["arrow"]:
                    c.itemconfigure(d["arrow"], text="▾")
            else:
                d["desc_id"] = c.create_text(
                    d["x"] + 12, d["y"] + d["desc_y"], anchor="nw",
                    text=d["dsc"], fill="#a8b2c0", font=self.f_chip,
                    width=tile_w - 36, tags=(d["tag"],))
                bb = c.bbox(d["desc_id"])
                dh = (bb[3] - bb[1]) + 8
                if d["arrow"]:
                    c.itemconfigure(d["arrow"], text="▴")
            # everything in this tile below the description slides
            thresh = d["y"] + d["desc_y"] - 2
            for it in c.find_withtag(d["tag"]):
                if it in (d["card"], d["desc_id"]):
                    continue
                bb2 = c.bbox(it)
                if bb2 and bb2[1] >= thresh:
                    c.move(it, 0, dh)
            x0, y0, x1, y1 = c.coords(d["card"])
            c.coords(d["card"], x0, y0, x1, y1 + dh)
            d["h"] += dh
            self._gal_relayout()


        def _gal_pick(self, pid):
            if self.tree.exists(pid):
                self.tree.selection_set(pid)
            for tpid, d in getattr(self, "_gtiles", {}).items():
                st = (self.rows.get(tpid) or {}).get("status")
                self.gal.itemconfigure(
                    d["card"],
                    outline=ACCENT if tpid == pid
                    else STATUS_COLOR.get(st, PANEL2))

        def _check_game_update(self):
            """After a scan, quietly compare the newest game binary's
            date against the newest date announced on run8studios.com.
            Updates rewrite the game EXE/DLLs, so an older local date
            with a newer site announcement means the updater is due --
            and it updates DLC files too."""
            if getattr(self, "_upd_checking", False):
                return
            if not load_json(DATA_DIR / "config.json",
                             {}).get("update_check", True):
                return
            if DEMO_MODE:
                # demo: always show what the alerts look like
                self._flag_game_update(datetime.now(), auto=False)
                self._flag_new_products(2)
                return
            self._upd_checking = True

            def work():
                verdict = None
                try:
                    root = Path(self.app().config.get("run8_install",
                                                      ""))
                    local = 0
                    if root.is_dir():
                        for f in root.glob("*"):
                            if (f.suffix.lower() in (".exe", ".dll")
                                    and "updater" not in f.name.lower()):
                                local = max(local, f.stat().st_mtime)
                    if local:
                        import re as _re
                        page = html_to_text(fetch(
                            "https://www.run8studios.com/", timeout=15))
                        mon = ("JANUARY FEBRUARY MARCH APRIL MAY JUNE "
                               "JULY AUGUST SEPTEMBER OCTOBER NOVEMBER "
                               "DECEMBER").split()
                        best = None
                        for m_ in _re.finditer(
                                r"\b(" + "|".join(mon) +
                                r")\s+(\d{1,2}),?\s+(20\d\d)",
                                page.upper()):
                            try:
                                d_ = datetime(
                                    int(m_.group(3)),
                                    mon.index(m_.group(1)) + 1,
                                    int(m_.group(2)))
                            except ValueError:
                                continue
                            if best is None or d_ > best:
                                best = d_
                        if best is not None:
                            loc = datetime.fromtimestamp(local)
                            verdict = (None if best.date()
                                       <= loc.date() else best)
                except Exception:
                    verdict = None
                newn = 0
                try:
                    # new releases often ride along with updates: scan
                    # the store pages for product links the catalog
                    # doesn't know yet (2 quick fetches, no per-product
                    # pass)
                    known = set()
                    for pr_ in self.app().products:
                        for u_ in ([pr_.get("url")]
                                   + list(pr_.get("alt_urls", []))):
                            if u_:
                                known.add(Path(urllib.parse.urlparse(
                                    u_).path).name.lower())
                    fresh = set()
                    for pu in self.app().config.get("catalog_pages",
                                                    []):
                        try:
                            h_ = fetch(pu, timeout=15)
                        except Exception:
                            continue
                        for slug in SLUG_RE.findall(h_):
                            if slug.lower() not in known:
                                fresh.add(slug.lower())
                    newn = len(fresh)
                except Exception:
                    newn = 0
                finally:
                    self._upd_checking = False
                    self.q.put(("call",
                                lambda v=verdict, n=newn:
                                (self._flag_game_update(v),
                                 self._flag_new_products(n))))
            threading.Thread(target=work, daemon=True).start()

        def _flag_game_update(self, site_date, auto=True):
            b = getattr(self, "_upd_btn", None)
            if site_date is None:
                if b is not None:
                    b.destroy()
                    self._upd_btn = None
                return
            if b is not None:
                return
            bar = self.toolbar_btns[0].master
            b = ttk.Button(bar, text="Game update available!",
                           style="Accent.TButton",
                           command=self._upd_click)
            b.pack(side="right", padx=(6, 0))
            self._upd_btn = b
            attach_tip(b, "The Run8 site announced an update newer than "
                          "your game files. Click to run the official "
                          "updater -- it updates DLC files too.")
            self.log("Run8 update available -- the site announcement "
                     f"({site_date:%Y-%m-%d}) is newer than your game "
                     "files."
                     + (" (demo preview)" if DEMO_MODE else ""), "warn")
            if auto and not DEMO_MODE:
                self._run_updater_stream()

        def restore_backup(self):
            if self.busy:
                return
            if DEMO_MODE:
                if not messagebox.askyesno(
                        "Restore from backup -- REPLACES YOUR SETUP",
                        "(demo -- simulated)\n\nRestore now?",
                        icon="warning", default="no"):
                    return
                self._demo_play(
                    ["backup holds 6 record file(s) and 67 installer "
                     "file(s) (3.20 GB)",
                     "current records snapshotted -> "
                     "pre_restore_20260711_190000",
                     "6 record file(s) restored",
                     "[dl] 25%  (17/67)", "[dl] 50%  (34/67)",
                     "[dl] 75%  (51/67)", "[dl] 100%  (67/67)",
                     "restore complete (demo -- nothing touched)"])
                return
            dflt = load_json(DATA_DIR / "config.json", {}
                             ).get("backup_dir") or \
                DEFAULT_CONFIG["backup_dir"]
            zp = filedialog.askopenfilename(
                parent=self, title="Pick the backup to restore",
                initialdir=dflt,
                filetypes=[("Run8DLC backup", "*.zip")])
            if not zp:
                return
            if not messagebox.askyesno(
                    "Restore from backup -- REPLACES YOUR SETUP",
                    "This OVERWRITES your current ledger, records, "
                    "settings and installers with the backup's "
                    "contents.\n\nYour current record files are "
                    "snapshotted to a pre_restore folder first, so "
                    "this can be undone -- but it is still a big "
                    "hammer.\n\nRestore now?",
                    icon="warning", default="no"):
                return
            self._run_cli(
                ["restore-backup", zp, "--yes"],
                done=lambda: (messagebox.showinfo(
                    "Restore complete",
                    "Restored. The manager restarts now to load "
                    "everything."), restart_app()))

        def _back_to_view(self):
            self._set_mode(getattr(self, "_last_view", "Gallery"))

        def _upd_click(self):
            if DEMO_MODE:
                self._demo_play(
                    ["running the official Run8 updater... (demo)",
                     "[updater] checking file manifest...",
                     "[updater] 14 files need updating",
                     "[dl] 40%", "[dl] 80%", "[dl] 100%",
                     "updater finished (ok) (demo)"])
                return
            self._run_updater_stream()

        def _flag_new_products(self, n):
            b = getattr(self, "_new_btn", None)
            if not n:
                if b is not None:
                    b.destroy()
                    self._new_btn = None
                return
            if b is not None:
                return
            bar = self.toolbar_btns[0].master
            b = ttk.Button(bar, text="New DLC in store!",
                           style="Accent.TButton",
                           command=self._fetch_new_products)
            b.pack(side="right", padx=(6, 0))
            self._new_btn = b
            attach_tip(b, "The store lists products this catalog "
                          "doesn't know yet. Click to add just the new "
                          "ones -- quick, no full price refresh.")
            self.log(f"{n} new product(s) spotted on the 3DTS store -- "
                     "click 'New DLC in store!' to add them"
                     + (" (demo preview)" if DEMO_MODE else ""), "warn")

        def _fetch_new_products(self):
            if self.busy:
                return
            b = getattr(self, "_new_btn", None)
            if b is not None:
                b.destroy()
                self._new_btn = None
            if DEMO_MODE:
                self._demo_play(
                    ["Scanning store pages ...",
                     "2 NEW products found in the store:",
                     "  + BNSF Scenic Sub  $34  (demo)",
                     "  + EMD SD70ACe Pack 2  $25  (demo)",
                     "downloading 2 new image(s)... done (demo)"])
                return
            self.log("adding new store products (quick scan -- only "
                     "the new ones are fetched)...", "accent")
            self._run_cli(["refresh"], done=lambda: self._run_cli(
                ["media"], done=self._media_done))

        def _run_updater_stream(self):
            """Run the official updater as a child process with its
            output flowing into the app log. If it can't be captured
            (elevation, GUI-only), fall back to launching it."""
            if getattr(self, "_upd_running", False):
                return
            self._upd_running = True
            app = self.app()
            exe = updater_path(app)
            if not exe.is_file():
                self._upd_running = False
                self.log("updater not found -- check the 'Run8 install' "
                         "folder in Settings", "warn")
                return
            self.log("running the official Run8 updater...", "accent")

            def work():
                try:
                    pr = subprocess.Popen(
                        [str(exe)], cwd=str(exe.parent),
                        stdout=subprocess.PIPE,
                        stderr=subprocess.STDOUT, text=True,
                        creationflags=CREATE_NO_WINDOW)
                    for line in pr.stdout:
                        line = line.rstrip()
                        if line:
                            self.log("  [updater] " + line)
                    rc = pr.wait()
                    self.log(f"updater finished "
                             f"({'ok' if rc == 0 else f'exit {rc}'})",
                             "accent" if rc == 0 else "warn")
                    if rc == 0:
                        self.q.put(("call", lambda:
                                    self._flag_game_update(None)))
                except Exception as e:
                    self.log(f"couldn't capture the updater ({e!r}) -- "
                             "opening it normally", "warn")
                    try:
                        launch(exe, cwd=exe.parent)
                    except Exception:
                        pass
                finally:
                    self._upd_running = False
            threading.Thread(target=work, daemon=True).start()

        def _media_done(self):
            self._imgs.clear()
            self._thumbs.clear()
            self._gal_key = None
            self._det_imgs = {}
            self._det_shown = None
            self.rescan()

        def _hero(self, pid):
            img = self._imgs.get(pid)
            if img is None:
                f = DATA_DIR / "media" / (pid + ".png")
                try:
                    img = tk.PhotoImage(file=str(f)) if f.exists() else False
                except Exception:
                    img = False
                self._imgs[pid] = img
            return img

        def _close_card(self, *_):
            c = getattr(self, "_card", None)
            if c is not None:
                try:
                    c.destroy()
                except Exception:
                    pass
                self._card = None

        def _demo_play(self, lines, done=None):
            """Demo mode: act out a workflow in the log with realistic
            pacing and progress-bar motion -- nothing touches disk."""
            self.q.put(("busy", True))

            def step(i=0):
                if i >= len(lines):
                    self.q.put(("progress", None))
                    self.q.put(("busy", False))
                    if done:
                        done()
                    return
                ln = lines[i]
                m = re.match(r"\[dl\]\s+(\d+)%", ln)
                if m:
                    self.q.put(("progress", int(m.group(1))))
                self.log("  " + ln)
                self.after(160, lambda: step(i + 1))
            step()

        def _demo_flip(self, pid, status):
            """Demo Disable/Enable: flip the fake status for real-feel
            testing; Check My Collection resets the whole demo."""
            r = self.rows.get(pid)
            if not r:
                return
            r["status"] = status
            if status == "quarantined":
                self.quar = set(self.quar) | {pid}
                self.log(f"{r['name']} disabled -- folders parked in "
                         "'uninstalled' (demo)", "accent")
            else:
                self.quar = set(self.quar) - {pid}
                self.log(f"{r['name']} enabled -- folders put back "
                         "(demo)", "accent")
            self._render(self.rows, self.state, self.quar, [])

        def _tile_buy(self, pid):
            try:
                if self.tree.exists(pid):
                    self.tree.selection_set(pid)
                self.open_store()
            except Exception as e:
                self.log(f"buy failed: {e!r}", "warn")

        def _card_act(self, pid, fn):
            self._close_card()
            try:
                if self.tree.exists(pid):
                    self.tree.selection_set(pid)
                fn()
            except Exception as e:
                self.log(f"action failed: {e!r}", "warn")

        def _autosize_columns(self, items):
            """Columns hug their longest cell (plus a little air), so the
            detail pane keeps the leftover width."""
            try:
                f = tkfont.nametofont("TkDefaultFont")
                fh = tkfont.Font(font=f)
                fh.configure(weight="bold")
                heads = {"name": "Product", "category": "Category",
                         "status": "Status", "price": "Price",
                         "evidence": "How I Know"}
                getv = {"name": lambda r: r["name"],
                        "category": lambda r: r["category"],
                        "status": lambda r: STATUS_LABEL[r["status"]],
                        "price": lambda r: ("" if r.get("price") is None
                                            else f"${r['price']:.2f}"),
                        "evidence": lambda r: (r.get("evtext") or "")[:60]}
                for c, head in heads.items():
                    w = fh.measure(head) + 38
                    for r in items:
                        w = max(w, f.measure(getv[c](r)) + 34)
                    w = min(w, 400 if c in ("name", "evidence") else 200)
                    self.tree.column(c, width=w, stretch=False)
            except Exception as e:
                self.log(f"column autosize failed: {e!r}", "warn")

        def _sort_by(self, col):
            if self.sort_col == col:
                self.sort_rev = not self.sort_rev
            else:
                self.sort_col, self.sort_rev = col, False
            self._refilter()

        # ------------------------------------------------------- detail panel

        def _sel(self):
            s = self.tree.selection()
            return self.rows.get(s[0]) if s else None

        def _on_select(self):
            r = self._sel()
            for b in (self.b_install, self.b_uninst, self.b_restore,
                      self.b_buy, self.b_store):
                b.configure(state="disabled")
            if not r:
                self.d_name.configure(text="Select a product")
                self.d_meta.configure(text="")
                self.d_status.configure(text="", fg=DIM)
                self.d_desc.configure(text="")
                self._set_ev("")
                return
            self.d_name.configure(text=r["name"])
            self.d_meta.configure(text=f"{r['category']}   ·   {price_str(r)}")
            self.d_status.configure(text=STATUS_LABEL[r["status"]],
                                    fg=STATUS_COLOR[r["status"]])
            d = (self.app().by_id.get(r["pid"], {}) or {}).get("desc") or ""
            self.d_desc.configure(
                text=d[:340] + ("…" if len(d) > 340 else ""))
            self._load_page(r)
            ev, lines = r["ev"], []
            for f in ev["installers"]:
                lines.append(f"Installer on disk: {f.name}")
            for f in ev["receipts"]:
                lines.append(f"Receipt: {f.name}")
            for g in ev["game"]:
                lines.append(f"In the game folder: {g}")
            for g in ev.get("game_files", []):
                lines.append(f"In the game folder: {g}")
            for it in ev["ledger"]:
                d = (it.get("date") or "")[:10]
                tx = it.get("transaction_id") or ""
                lines.append(f"Purchase record: {d} {tx}".rstrip())
            if r["pid"] in self.quar:
                lines.append("Disabled -- folders parked in the "
                             "'uninstalled' folder")
            if r["status"] == "owned":
                lines.append("")
                lines.append("Owned, but not installed in the game folder")
            if not lines:
                lines = ["No proof of purchase on file yet"]
            self._set_ev("\n".join(lines))

            if not self.busy:
                app = self.app()
                exe = find_installer_for(app, r["pid"])
                r["_exe"] = exe
                if exe:
                    self.b_install.configure(
                        state="normal",
                        text="Reinstall" if r["status"] in
                             ("installed", "quarantined") else "Run installer")
                good, _sk = _game_paths_for(app, r["pid"], self.state)
                if good or (DEMO_MODE
                            and r["status"] == "installed"):
                    self.b_uninst.configure(state="normal")
                if r["pid"] in self.quar:
                    self.b_restore.configure(state="normal")
                if r.get("url"):
                    (self.b_buy if r["status"] == "missing"
                     else self.b_store).configure(state="normal")


        def _load_page(self, r):
            pid = r["pid"]
            self._det_r = r
            img = self._hero(pid)
            if not img:
                self._det_shown = None
                self.d_img.configure(
                    image="", text="no image yet --\n"
                                   "Settings -> Refresh Store Prices")
                return
            # scale the photo up (rationally) to use the panel's width
            avail = max(self.det.winfo_width() - 40, 200)
            iw = max(img.width(), 1)
            best = (1, 1)
            for a_, b_ in ((3, 1), (5, 2), (2, 1), (3, 2), (1, 1)):
                if iw * a_ / b_ <= avail:
                    best = (a_, b_)
                    break
            if (pid, best) == getattr(self, "_det_shown", None):
                return
            if not hasattr(self, "_det_imgs"):
                self._det_imgs = {}
            t = self._det_imgs.get((pid, best))
            if t is None:
                try:
                    a_, b_ = best
                    t = img.zoom(a_, a_) if a_ > 1 else img
                    if b_ > 1:
                        t = t.subsample(b_, b_)
                except Exception:
                    t = img
                if len(self._det_imgs) > 12:
                    self._det_imgs.clear()
                self._det_imgs[(pid, best)] = t
            self._det_shown = (pid, best)
            self.d_img.configure(image=t, text="")


        def _set_ev(self, text):
            self.d_ev.configure(state="normal")
            self.d_ev.delete("1.0", "end")
            self.d_ev.insert("1.0", text)
            self.d_ev.configure(state="disabled")

        # ------------------------------------------------------------ actions

        def open_store(self):
            r = self._sel()
            if r and r.get("url"):
                webbrowser.open(r["url"])
                self.log(f"opened store page for {r['name']}")

        def run_updater(self):
            app = self.app()
            exe = updater_path(app)
            if not exe.is_file():
                messagebox.showerror("Run8 Updater",
                                     f"Updater not found:\n{exe}\n\n"
                                     "Open the Settings tab and check "
                                     "that the 'Run8 install' folder "
                                     "points at your game.")
                return
            launch(exe, cwd=exe.parent)
            self.log(f"launched {exe.name}", "accent")

        def reinstall(self):
            r = self._sel()
            if not r:
                return
            if DEMO_MODE:
                self._demo_play(
                    [f"launched installer: r8v3_{r['pid']}.exe (demo)"])
                return
            if not r.get("_exe"):
                return
            launch(r["_exe"], cwd=r["_exe"].parent)
            self.log(f"launched installer: {r['_exe'].name}", "accent")

        def uninstall(self):
            r = self._sel()
            if not r:
                return
            if DEMO_MODE:
                if messagebox.askyesno(
                        "Disable",
                        f"Disable {r['name']}?\n\n(demo -- the "
                        "status flips; nothing on disk changes)"):
                    self._demo_flip(r["pid"], "quarantined")
                return
            app = self.app()
            good, skipped = _game_paths_for(app, r["pid"], self.state)
            if not good:
                messagebox.showinfo(
                    "Disable",
                    "This product has no folders of its own inside the "
                    "game -- equipment packs share files with other "
                    "products, so only routes can be disabled "
                    "individually.")
                return
            msg = (f"Disable {r['name']}?\n\nThese folders will be MOVED "
                   f"(not deleted) into the tool's 'uninstalled' folder:\n\n"
                   + "\n".join(f"  {p}" for p in good)
                   + "\n\nRestore puts them straight back.")
            if not messagebox.askyesno("Disable", msg,
                                       icon="warning"):
                return
            self.q.put(("busy", True))

            def work():
                try:
                    ok, lines = uninstall_product(app, r["pid"], apply=True,
                                                       state=self.state)
                    for ln in lines:
                        self.log(ln, "accent" if ok else "warn")
                finally:
                    self.q.put(("busy", False))
                    self.q.put(("call", self.rescan))

            threading.Thread(target=work, daemon=True).start()

        def restore(self):
            r = self._sel()
            if not r:
                return
            if DEMO_MODE:
                self._demo_flip(r["pid"], "installed")
                return
            self.q.put(("busy", True))

            def work():
                try:
                    ok, lines = restore_product(self.app(), r["pid"])
                    for ln in lines:
                        self.log(ln, "accent" if ok else "warn")
                finally:
                    self.q.put(("busy", False))
                    self.q.put(("call", self.rescan))

            threading.Thread(target=work, daemon=True).start()

        def show_transactions(self):
            self._set_mode("Transactions")

        def _pane_transactions(self, f):
            app = self.app()
            items = sorted(app.ledger.get("items", []),
                           key=lambda i: (i.get("date") or "",
                                          i.get("name") or ""),
                           reverse=True)
            head = tk.Frame(f, bg=PANEL)
            head.pack(fill="x")
            tk.Label(head, text="Transaction history", bg=PANEL,
                     fg=ACCENT, font=self.f_hdr).pack(side="left")
            ttk.Button(head, text="Back",
                       command=self._back_to_view).pack(side="right")
            cols = ("date", "product", "txid", "exe", "source")
            tv = ttk.Treeview(f, columns=cols, show="headings")
            widths = {"date": 90, "product": 260, "txid": 170,
                      "exe": 220, "source": 80}
            for c_ in cols:
                tv.heading(c_, text=c_.title())
                tv.column(c_, width=widths[c_], anchor="w",
                          stretch=(c_ in ("product", "exe")))
            for i in items:
                tv.insert("", "end", values=(
                    i.get("date") or "?", i.get("name") or "?",
                    i.get("transaction_id") or i.get("txid") or "-",
                    i.get("exe") or "-", i.get("source") or "-"))
            tv.pack(fill="both", expand=True, pady=(10, 6))
            pth = DATA_DIR / "transactions.txt"
            tk.Label(f, text=f"{len(items)} purchases   ·   flat "
                             f"file: {pth}",
                     bg=PANEL, fg=DIM).pack(anchor="w", pady=(0, 4))

        def import_records_dialog(self):
            self._set_mode("Import")

        def _pane_import(self, f):
            """Guided import: explains each record source (including how
            to export emails) before any file picker appears."""
            head = tk.Frame(f, bg=PANEL)
            head.pack(fill="x")
            tk.Label(head, text="Import purchase records", bg=PANEL,
                     fg=ACCENT, font=self.f_hdr).pack(side="left")
            ttk.Button(head, text="Back",
                       command=self._back_to_view).pack(side="right")
            ttk.Label(f, style="Dim.TLabel", wraplength=620,
                      justify="left",
                      text="Pick whichever matches how your records "
                           "are kept. Every transaction ID found is "
                           "added to the ledger; nothing is ever "
                           "overwritten."
                      ).pack(anchor="w", pady=(4, 10))
            CARDS = (
                ("Receipt screenshots",
                 "PNG / JPG pictures of receipts or the store's "
                 "download page. Windows' built-in OCR reads the "
                 "transaction IDs off them.",
                 "screens"),
                ("Emails from the store",
                 "First save each receipt email as a file:\n"
                 "  • Gmail: open the email, click the \u22ee menu "
                 "at the top right of the message, choose "
                 "'Download message'.\n"
                 "  • Outlook: drag the email from the list into "
                 "any folder.\n"
                 "  • Apple Mail: drag the message onto the "
                 "Desktop.\n"
                 "Each becomes a .eml file -- pick those below.",
                 "emails"),
                ("A document or spreadsheet",
                 "txt, csv, Word or Excel. Each purchase just needs "
                 "the product and its transaction ID on the same or "
                 "nearby lines.",
                 "docs"),
            )
            for title, blurb, kind in CARDS:
                card = tk.Frame(f, bg=PANEL2, padx=12, pady=10)
                card.pack(fill="x", pady=(0, 8))
                top = tk.Frame(card, bg=PANEL2)
                top.pack(fill="x")
                tk.Label(top, text=title, bg=PANEL2, fg=FG,
                         font=self.f_big).pack(side="left")
                ttk.Button(top, text="Choose files…",
                           style="Accent.TButton",
                           command=lambda k=kind:
                           (self._back_to_view(),
                            self._pick_and_import(k))
                           ).pack(side="right")
                tk.Label(card, text=blurb, bg=PANEL2, fg=DIM,
                         wraplength=560, justify="left"
                         ).pack(anchor="w", pady=(4, 0))
            foot = tk.Frame(f, bg=PANEL)
            foot.pack(fill="x", pady=(4, 0))
            ttk.Button(foot, text="Pick any mix of files",
                       command=lambda: (self._back_to_view(),
                                        self._pick_and_import("any"))
                       ).pack(side="left")

        def _pick_and_import(self, kind):
            try:
                ft = {
                    "screens": [("Receipt screenshots",
                                 "*.png *.jpg *.jpeg *.bmp")],
                    "emails": [("Email receipts", "*.eml")],
                    "docs": [("Documents", "*.txt *.csv *.log *.md "
                              "*.docx *.xlsx")],
                }.get(kind, [("All supported",
                              "*.png *.jpg *.jpeg *.bmp *.txt *.csv "
                              "*.log *.md *.docx *.xlsx *.eml")])
                titles = {"screens": "Pick your receipt screenshots",
                          "emails": "Pick your saved .eml receipt emails",
                          "docs": "Pick your record documents"}
                picks = filedialog.askopenfilenames(
                    parent=self,
                    title=titles.get(kind, "Pick purchase record files"),
                    filetypes=ft)
                if picks:
                    self._import_files(picks)
            except Exception as e:
                self.log(f"import failed: {e!r}", "warn")

        def _import_files(self, picks):
            try:
                imgs = [p for p in picks if Path(p).suffix.lower()
                        in (".png", ".jpg", ".jpeg", ".bmp")]
                docs = [p for p in picks
                        if Path(p).suffix.lower() in RECORD_EXTS]
                cfg = load_json(DATA_DIR / "config.json", {})
                tx_dir = Path(cfg.get("transactions_dir", ""))
                if imgs and str(tx_dir):
                    tx_dir.mkdir(parents=True, exist_ok=True)
                    for p in imgs:
                        dst = tx_dir / Path(p).name
                        if Path(p).resolve() != dst.resolve():
                            shutil.copy2(p, dst)
                if docs and imgs:
                    self._run_cli(["import-records"] + list(docs),
                                  done=lambda: self._run_cli(
                                      ["ocr-receipts"], done=self.rescan))
                elif docs:
                    self._run_cli(["import-records"] + list(docs),
                                  done=self.rescan)
                elif imgs:
                    self._run_cli(["ocr-receipts"], done=self.rescan)
                else:
                    self.log("no supported files in that selection "
                             "(need images, .eml, or documents)", "warn")
            except Exception as e:
                self.log(f"import failed: {e!r}", "warn")

        def _when_idle(self, fn, tries=100):
            if self.busy and tries > 0:
                self.after(300, lambda: self._when_idle(fn, tries - 1))
            else:
                fn()

        def backup_now(self):
            """One click: pack installers + records into the configured
            Backups folder. Progress shows on the bar; the zip replaces
            the previous one atomically."""
            if self.busy:
                return
            if DEMO_MODE:
                self._demo_play(
                    ["packing 67 installer file(s)...",
                     "[dl] 25%  (17/67)", "[dl] 50%  (34/67)",
                     "[dl] 75%  (51/67)", "[dl] 100%  (67/67)",
                     "backup complete: 67 installer file(s) + records "
                     "packed -- 3.20 GB in, 3.05 GB zip (demo -- "
                     "nothing written)"])
                return
            dest = load_json(DATA_DIR / "config.json", {}
                             ).get("backup_dir") or \
                DEFAULT_CONFIG["backup_dir"]
            self.log(f"backing up to {dest} ...", "accent")
            self._run_cli(["backup"])

        def purge_quarantine(self):
            if self.busy:
                return
            if DEMO_MODE:
                if messagebox.askyesno(
                        "Permanently delete disabled items",
                        "Delete them forever? (demo -- simulated)",
                        icon="warning", default="no"):
                    self._demo_play(
                        ["deleted 2 folder(s), 512 MB freed (demo)"])
                return
            try:
                n_files, total = quarantine_size(self.app())
                quar = load_json(DATA_DIR / "quarantine.json",
                                 {"items": []})
                n_items = len(quar.get("items", []))
                if not n_files and not n_items:
                    messagebox.showinfo("Permanently delete disabled items",
                                        "There are no disabled items -- "
                                        "nothing to delete.")
                    return
                if not messagebox.askyesno(
                        "Permanently delete disabled items",
                        f"Disabled items on hand: {n_items} product(s) "
                        f"({total/1048576:.0f} MB).\n\n"
                        "This DELETES the saved files permanently. "
                        "Restore will no longer be possible -- getting "
                        "them back would mean running the installers "
                        "again.\n\nDelete them forever?",
                        icon="warning", default="no"):
                    return
                self._run_cli(["purge-quarantine", "--yes"],
                              done=self.rescan)
            except Exception as e:
                self.log(f"purge failed: {e!r}", "warn")

        def show_settings(self):
            self._set_mode("Settings")

        # ----------------------------------------------------- subprocess ops

        def _run_cli(self, cli_args, done=None):
            """Stream a run8dlc.py subcommand's output into the log."""
            self.q.put(("busy", True))
            if getattr(sys, "frozen", False):
                cmd = [sys.executable]
            else:
                cmd = [sys.executable, "-u",
                       str(Path(__file__).resolve())]
            if self.config_path:
                cmd += ["--config", str(self.config_path)]
            cmd += cli_args

            def work():
                try:
                    p = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                         stderr=subprocess.STDOUT, text=True,
                                         bufsize=1, cwd=str(DATA_DIR),
                                         creationflags=CREATE_NO_WINDOW,
                                         env={**os.environ,
                                              "PYTHONUNBUFFERED": "1"})
                    self._proc = p
                    import re as _re
                    for line in p.stdout:
                        line = line.rstrip()
                        if not line.strip():
                            continue
                        m = _re.match(r"\[dl\]\s+(\d+)%", line)
                        if m:
                            self.q.put(("progress", int(m.group(1))))
                        self.log("  " + line)
                    rc = p.wait()
                    self.log(f"[{' '.join(cli_args)}] finished "
                             f"({'ok' if rc == 0 else f'exit {rc}'})",
                             "accent" if rc == 0 else "warn")
                except Exception as e:
                    self.log(f"failed: {e!r}", "warn")
                finally:
                    self.q.put(("progress", None))
                    self.q.put(("busy", False))
                    if done:
                        self.q.put(("call", done))

            threading.Thread(target=work, daemon=True).start()

        def update_store(self):
            if self.busy:
                return
            if DEMO_MODE:
                if messagebox.askyesno(
                        "Refresh Store Prices",
                        "Fetch current prices and images? "
                        "(demo -- simulated)"):
                    self._demo_play(
                        ["Scanning https://www.run8studios.com/"
                         "routes.shtml ...",
                         "Scanning https://www.run8studios.com/"
                         "trainsets.shtml ...",
                         "No new products since the bundled catalog.",
                         "[dl] 30%  UP Fresno Sub South",
                         "[dl] 65%  GE Dash 9 Pack 2",
                         "[dl] 100%  Wellcar Pack 3",
                         "Catalog saved (72 entries) (demo)"])
                return
            if not messagebox.askyesno(
                    "Refresh Store Prices",
                    "Fetch current prices, any newly released products, "
                    "and the product image library from the store?\n\n"
                    "Takes 2-3 minutes. Only needed at first setup or "
                    "when Run8 releases something new -- they don't do "
                    "sales."):
                return
            self._run_cli(["refresh", "--prices"],
                          done=lambda: self._run_cli(
                              ["media"], done=self._media_done))

        def open_report(self):
            if self.busy:
                return

            def done():
                p = DATA_DIR / "report.html"
                if p.exists():
                    webbrowser.open(p.as_uri())
            self._run_cli(["report"], done=done)

        def add_purchase(self):
            if self.busy:
                return
            self._set_mode("Add")

        def _pane_add_purchase(self, f):
            head = tk.Frame(f, bg=PANEL)
            head.pack(fill="x")
            tk.Label(head, text="Add a purchase", bg=PANEL, fg=ACCENT,
                     font=self.f_hdr).pack(side="left")
            ttk.Button(head, text="Back",
                       command=self._back_to_view).pack(side="right")
            v_mode = tk.StringVar(value="url")
            v_url, v_tx = tk.StringVar(), tk.StringVar()
            v_file = tk.StringVar()
            hint = ttk.Label(f, style="Dim.TLabel", wraplength=600,
                             justify="left")
            frm = ttk.Frame(f, style="Panel.TFrame")
            frm.columnconfigure(1, weight=1)

            def _pick(var):
                p_ = filedialog.askopenfilename(
                    parent=self,
                    title="Pick the installer / email / add-on file",
                    filetypes=[("Installer EXE", "*.exe"),
                               ("Purchase email", "*.eml"),
                               ("Any file", "*.*")])
                if p_:
                    var.set(p_)

            def _mkrow(label, var, browse=False):
                lab = ttk.Label(frm, text=label, style="Panel.TLabel")
                ent = ttk.Entry(frm, textvariable=var, width=54)
                wids = [lab, ent]
                if browse:
                    wids.append(ttk.Button(frm, text="Browse…",
                                           width=9,
                                           command=lambda: _pick(var)))
                return wids

            rows = {"url": _mkrow("Receipt page", v_url),
                    "txid": _mkrow("Transaction ID", v_tx),
                    "file": _mkrow("Installer file", v_file,
                                   browse=True)}
            HINTS = {
                "url": "Copy the address bar on the 'Transaction "
                       "Approved' page while it is still open -- "
                       "everything else is automatic. (If it was on "
                       "your clipboard, it is already filled in.)",
                "txid": "Paste the ID itself -- or a whole download "
                        "link, and the ID is pulled out automatically. "
                        "The manager then downloads the installer and "
                        "files everything.",
                "file": "An installer (or any add-on file) is copied "
                        "into your Installers folder and recorded -- "
                        "add the transaction ID too if you have it. A "
                        "saved purchase email (.eml) is read for its "
                        "transaction details automatically."}
            NEED = {"url": ("url",), "txid": ("txid",),
                    "file": ("file",)}
            SHOW = {"url": ("url",), "txid": ("txid",),
                    "file": ("file", "txid")}

            def _update(*_):
                m = v_mode.get()
                hint.configure(text=HINTS[m])
                for wids in rows.values():
                    for w_ in wids:
                        w_.grid_remove()
                for rr, key in enumerate(SHOW[m]):
                    wids = rows[key]
                    wids[0].grid(row=rr, column=0, sticky="w",
                                 padx=(0, 8), pady=2)
                    wids[1].grid(row=rr, column=1, sticky="we",
                                 pady=2)
                    if len(wids) > 2:
                        wids[2].grid(row=rr, column=2, sticky="w",
                                     padx=(6, 0), pady=2)

            for val, lbl in (("url",
                              "I have the receipt page address "
                              "(easiest)"),
                             ("txid",
                              "I only have the transaction ID"),
                             ("file",
                              "I have the installer file or a "
                              "purchase email")):
                tk.Radiobutton(f, text=lbl, variable=v_mode,
                               value=val, bg=PANEL, fg=FG,
                               selectcolor=FIELD,
                               activebackground=PANEL,
                               activeforeground=FG, command=_update
                               ).pack(anchor="w", pady=(2, 0))
            hint.pack(anchor="w", pady=(6, 8))
            frm.pack(fill="x")
            v_inst = tk.BooleanVar(value=True)
            tk.Checkbutton(f,
                           text="Launch the installer when it's here",
                           variable=v_inst, bg=PANEL, fg=FG,
                           selectcolor=FIELD, activebackground=PANEL,
                           activeforeground=FG
                           ).pack(anchor="w", pady=(10, 12))
            row = tk.Frame(f, bg=PANEL)
            row.pack(fill="x")

            def go(*_):
                m = v_mode.get()
                vals = {"url": v_url.get().strip(),
                        "txid": v_tx.get().strip(),
                        "file": v_file.get().strip()}
                for need in NEED[m]:
                    if not vals[need]:
                        messagebox.showwarning(
                            "Add purchase", "Please fill in every "
                            "field for this option.", parent=self)
                        return
                if m == "url" and not vals["url"].lower(
                        ).startswith("http"):
                    messagebox.showwarning("Add purchase",
                                           "That doesn't look like a "
                                           "URL.", parent=self)
                    return
                if m == "txid" and ("/" in vals["txid"]
                                    or vals["txid"].lower(
                                        ).startswith("http")):
                    t_ = txid_from_url(vals["txid"])
                    if not t_:
                        messagebox.showwarning(
                            "Add purchase", "Couldn't find a "
                            "transaction ID in that link.",
                            parent=self)
                        return
                    vals["txid"] = t_
                if DEMO_MODE:
                    tx = vals["txid"] or "4X9DEMO812K"
                    prod = next((rr["name"] for rr in
                                 self.rows.values()
                                 if rr["status"] == "missing"),
                                "EMD MP15 Pack 3")
                    self._back_to_view()
                    self._demo_play(
                        [f"Transaction ID: {tx}",
                         "Download link: http://www.run8-services"
                         f".com/download.php?transid={tx}",
                         "[dl] 20%", "[dl] 45%", "[dl] 70%",
                         "[dl] 100%",
                         f"Product: {prod}",
                         "Logged to ledger + transactions.txt "
                         "(demo -- nothing written)"])
                    return
                if m == "file" and vals["file"].lower(
                        ).endswith(".eml"):
                    args = ["import-records", vals["file"]]
                else:
                    args = ["add"]
                    if m == "url":
                        args.append(vals["url"])
                    if vals["txid"] and m != "url":
                        args += ["--txid", vals["txid"]]
                    if m == "file":
                        args += ["--exe", vals["file"]]
                    if v_inst.get():
                        args.append("--install")
                self._back_to_view()
                self._run_cli(args, done=self.rescan)

            ttk.Button(row, text="Start", style="Accent.TButton",
                       command=go).pack(side="right")
            for wids in rows.values():
                wids[1].bind("<Return>", go)
            try:
                clip = self.clipboard_get().strip()
                if clip.lower().startswith("http") and len(clip) < 500:
                    v_url.set(clip)
            except Exception:
                pass
            _update()

    LOCO_SCHEMES = {
        "Boxcar Slate":     dict(type="GP40-2", kind="hood", axles=4,
                                 fans=3, high_nose=True, body="#37b6a7",
                                 stripe="#10151d", nose="#2a8f84"),
        "Amtrak Phase III": dict(type="P42", kind="genesis",
                                 body="#c9ced6",
                                 bands=("#d5212e", "#f2f4f7", "#3d6fd6"),
                                 nose="#3a3f47"),
        "UP Armour Yellow": dict(type="SD70ACe", kind="hood", axles=6,
                                 wide_nose=True, flare=True, flag=True,
                                 body="#ffb612", stripe="#da291c",
                                 roof="#9aa0a4"),
        "ATSF Warbonnet":   dict(type="Dash 9", kind="hood", axles=6,
                                 wide_nose=True, body="#c7c9c7",
                                 stripe="#ffc72c", bonnet="#c8102e"),
        "BNSF Heritage II": dict(type="Dash 9", kind="hood", axles=6,
                                 wide_nose=True, body="#ff6720",
                                 stripe="#ffcd00", hood2="#26352a"),
        "SP Daylight":      dict(type="SD45-2", kind="hood", axles=6,
                                 long45=True, body="#e86a1f",
                                 stripe="#c8102e"),
        "Conrail Blue":     dict(type="GP38-2", kind="hood", axles=4,
                                 fans=2, body="#0079c1",
                                 stripe="#e8edf2"),
        "CSX Blue & Gold":  dict(type="MP15DC", kind="hood", axles=4,
                                 switcher=True, body="#2d5f9e",
                                 stripe="#fdb813", nose="#fdb813"),
        "Rio Grande Gold":  dict(type="SD40T-2", kind="hood", axles=6,
                                 tunnel=True, body="#2a2c30",
                                 stripe="#ffb81c"),
        "Chessie Yellow":   dict(type="GP40-2", kind="hood", axles=4,
                                 fans=3, body="#ffc425",
                                 stripe="#e8542f", roof="#1b365d"),
        "BN Cascade Green": dict(type="SD40-2", kind="hood", axles=6,
                                 body="#008249", stripe="#e8f0ea"),
        "High Contrast":    dict(type="GP38-2", kind="hood", axles=4,
                                 fans=2, body="#1a1a1a",
                                 stripe="#00e5ff", outline="#ffffff"),
    }

    # theme -> (car_kind, colors) x3; title splits across the cars
    CONSISTS = {
        "Boxcar Slate":     [("box", "#6b4f3f", "#5d4436", "#e8e2d8"),
                             ("gon", "#4a5560", "#3c454e", "#e8e2d8"),
                             ("box", "#7a4438", "#68392f", "#e8e2d8")],
        "Amtrak Phase III": [("amfleet", "#c9ced6", "#d5212e", "#12141a"),
                             ("amfleet", "#c9ced6", "#d5212e", "#12141a"),
                             ("amfleet", "#c9ced6", "#d5212e",
                              "#12141a")],
        "UP Armour Yellow": [("rack", "#d8b25a", "#b99742", "#3a2c10"),
                             ("cov2", "#9aa0a4", "#7f858a", "#22262a"),
                             ("box", "#6b4f3f", "#5d4436", "#e8e2d8")],
        "ATSF Warbonnet":   [("well", "#8a5a3a", "#c8102e", "#f4f6f9"),
                             ("well", "#8a5a3a", "#2d6a4f", "#f4f6f9"),
                             ("well", "#8a5a3a", "#3d5a8a", "#f4f6f9")],
        "BNSF Heritage II": [("well", "#8a5a3a", "#a34a2a", "#f4f6f9"),
                             ("well", "#8a5a3a", "#6a6e52", "#f4f6f9"),
                             ("well", "#8a5a3a", "#54687a", "#f4f6f9")],
        "SP Daylight":      [("reef", "#e8dfc8", "#c8102e", "#8a1616"),
                             ("reef", "#e8dfc8", "#c8102e", "#8a1616"),
                             ("reef", "#e8dfc8", "#c8102e", "#8a1616")],
        "Conrail Blue":     [("box", "#0079c1", "#0068a8", "#f0f4f8"),
                             ("cov2", "#9aa0a4", "#7f858a", "#22262a"),
                             ("box", "#6b4f3f", "#5d4436", "#e8e2d8")],
        "CSX Blue & Gold":  [("hop", "#3a3f45", "#2c3036", "#fdb813"),
                             ("hop", "#3a3f45", "#2c3036", "#fdb813"),
                             ("hop", "#3a3f45", "#2c3036", "#fdb813")],
        "Rio Grande Gold":  [("hop", "#4a4438", "#3a352c", "#ffb81c"),
                             ("box", "#26282c", "#1f2124", "#ffb81c"),
                             ("gon", "#4a5560", "#3c454e", "#ffb81c")],
        "Chessie Yellow":   [("cov2", "#ffc425", "#e8a815", "#1b365d"),
                             ("cov2", "#ffc425", "#e8a815", "#1b365d"),
                             ("box", "#ffc425", "#e8542f", "#1b365d")],
        "BN Cascade Green": [("box", "#00563f", "#004a36", "#e8f0ea"),
                             ("hop", "#3a3f45", "#2c3036", "#e8f0ea"),
                             ("box", "#00563f", "#004a36", "#e8f0ea")],
        "High Contrast":    [("box", "#0d0d0d", "#1a1a1a", "#00e5ff"),
                             ("box", "#0d0d0d", "#1a1a1a", "#00e5ff"),
                             ("box", "#0d0d0d", "#1a1a1a", "#00e5ff")],
    }

    def draw_title_train(cv, theme, font, scale=1.0):
        """Masthead consist. Vehicle heights follow the real spec sheet,
        relatively: autoracks (~20') over boxcars/reefers (~15-17') even
        with road-diesel rooflines (15'5"-15'11"), then covered hoppers
        and the Viewliner, the low P42 (14'4"), the lower Amfleet
        (12'8"), single-stacked containers, open hoppers, and finally
        gondolas (~9')."""
        s = scale
        L = lambda v: int(v * s)
        cv.delete("all")
        cv.configure(width=L(396), height=L(42), bg=BG,
                     highlightthickness=0)
        sc = LOCO_SCHEMES.get(theme, LOCO_SCHEMES["Boxcar Slate"])
        cars = CONSISTS.get(theme, CONSISTS["Boxcar Slate"])
        rail_y = L(38)
        cv.create_line(0, rail_y, L(396), rail_y, fill=DIM, width=L(2))
        for tx in range(4, 396, 14):
            cv.create_line(L(tx), rail_y, L(tx + 6), rail_y,
                           fill=FIELD, width=L(4))

        deck = rail_y - L(10)
        # rooftop y per vehicle class (smaller = taller), from real specs
        ROOF_LOCO, ROOF_GENESIS, ROOF_SWITCH = L(9), L(12), L(11)
        ROOF_BOX, ROOF_RACK, ROOF_COV2 = L(9), L(6), L(10)
        ROOF_VIEW, ROOF_AMF, ROOF_WELL = L(10), L(13), L(13)
        ROOF_HOP, ROOF_GON = L(13), L(19)

        def wheels(pxs):
            for px in pxs:
                cv.create_oval(px - L(4), rail_y - L(8), px + L(4),
                               rail_y, fill="#15171b", outline=DIM)

        def underframe(u0, u1):
            cv.create_rectangle(u0 + L(2), deck, u1 - L(2), deck + L(2),
                                fill="#15171b", outline="")

        def fuel_tank(u0, u1, color, w=26):
            cx0 = (u0 + u1) // 2 - L(w // 2)
            cv.create_rectangle(cx0, deck + L(2), cx0 + L(w),
                                deck + L(7), fill=color, outline="")
            cv.create_line(cx0 + L(2), deck + L(3), cx0 + L(w - 2),
                           deck + L(3),
                           fill=_mixc(color, "#ffffff", 0.25))

        def cab_window(w0, wt, w1, wb):
            cv.create_rectangle(w0, wt, w1, wb, fill="#bfd6e8",
                                outline="")
            cv.create_line((w0 + w1) // 2, wt, (w0 + w1) // 2, wb,
                           fill="#5a6a78")   # split-pane mullion

        def fit_font(text, maxw):
            f = tkfont.Font(font=font)
            f.configure(weight="bold")
            while f.measure(text) > maxw and f.cget("size") > 7:
                f.configure(size=f.cget("size") - 1)
            return f

        def readable(fg, bg):
            if abs(_lumc(fg) - _lumc(bg)) >= 0.32:
                return fg
            return "#14161a" if _lumc(bg) > 0.5 else "#f2f4f6"

        # ------------------------------------------------ locomotive
        x0 = L(4)
        tank_c = sc.get("tank", _mixc(sc["body"], "#000000", 0.35))
        if sc["kind"] == "genesis":
            x1 = x0 + L(104)
            top = ROOF_GENESIS
            cv.create_polygon(
                x0 + L(2), deck, x0 + L(2), top + L(12),
                x0 + L(15), top, x1 - L(3), top,
                x1, top + L(6), x1, deck,
                fill=sc["body"], outline="")
            cv.create_polygon(
                x0 + L(2), top + L(12), x0 + L(15), top,
                x0 + L(23), top, x0 + L(9), deck, x0 + L(2), deck,
                fill=sc.get("nose", "#3a3f47"), outline="")
            for i, bc in enumerate(sc.get("bands", ())):
                by = top + L(8 + i * 3)
                cv.create_rectangle(x0 + L(13), by, x1, by + L(2),
                                    fill=bc, outline="")
            cab_window(x0 + L(17), top + L(3), x0 + L(29), top + L(9))
            # monocoque skirt hides most of the underbody
            cv.create_rectangle(x0 + L(9), deck + L(2), x1 - L(2),
                                deck + L(5),
                                fill=sc.get("nose", "#3a3f47"),
                                outline="")
            lw = (x0 + L(16), x0 + L(27), x1 - L(27), x1 - L(16))
        elif sc.get("switcher"):
            x1 = x0 + L(78)
            top = ROOF_SWITCH
            hood_t = top + L(6)
            cab0 = x1 - L(20)
            cv.create_rectangle(x0 + L(3), hood_t, cab0, deck,
                                fill=sc["body"], outline="")
            cv.create_rectangle(x0, hood_t + L(3), x0 + L(3), deck,
                                fill=sc.get("nose", sc["body"]),
                                outline="")
            cv.create_rectangle(cab0, top, x1, deck, fill=sc["body"],
                                outline="")
            cab_window(cab0 + L(3), top + L(2), x1 - L(3), top + L(9))
            cv.create_rectangle(x0 + L(3), deck - L(3), x1, deck - L(1),
                                fill=sc.get("stripe", FG), outline="")
            cv.create_line(x0 + L(3), hood_t - L(2), cab0,
                           hood_t - L(2), fill=DIM)
            cv.create_rectangle(x0 + L(9), hood_t - L(3), x0 + L(12),
                                hood_t, fill="#15171b", outline="")
            fuel_tank(x0, x1, tank_c, w=18)
            lw = (x0 + L(12), x0 + L(24), x1 - L(24), x1 - L(12))
        else:
            x1 = x0 + L(104)
            top = ROOF_LOCO
            hood_t = top
            cab_t = top
            nose_w = L(11) if sc.get("wide_nose") else L(8)
            pilot = L(3)
            cab0 = x0 + pilot + nose_w
            cab1 = cab0 + L(15)
            nose_t = top if sc.get("high_nose") else top + L(7)
            cv.create_polygon(x0, deck + L(2), x0 + pilot, deck - L(4),
                              x0 + pilot, deck + L(2),
                              fill="#15171b", outline="")
            cv.create_rectangle(x0 + pilot, nose_t, cab0, deck,
                                fill=sc.get("bonnet",
                                            sc.get("nose", sc["body"])),
                                outline="")
            if sc.get("high_nose"):
                cv.create_line(cab0, nose_t, cab0, deck,
                               fill=_mixc(sc["body"], "#000000", 0.35))
            cv.create_rectangle(cab0, cab_t, cab1, deck,
                                fill=sc.get("bonnet", sc["body"]),
                                outline="")
            cv.create_rectangle(cab1, hood_t, x1 - L(2), deck,
                                fill=sc.get("hood2", sc["body"]),
                                outline="")
            if sc.get("hood2"):
                # Heritage-style duotone: colored top, body-orange lower,
                # separation stripe riding the hood shoulder
                hmid = hood_t + (deck - hood_t) * 2 // 5
                cv.create_rectangle(cab1, hmid + L(2), x1 - L(2), deck,
                                    fill=sc["body"], outline="")
                cv.create_rectangle(cab1, hmid, x1 - L(2), hmid + L(2),
                                    fill=sc.get("stripe", FG),
                                    outline="")
            if sc.get("bonnet"):
                cv.create_polygon(cab1, cab_t, cab1 + L(16), deck,
                                  cab1, deck,
                                  fill=sc["bonnet"], outline="")
                cv.create_line(cab1, cab_t, cab1 + L(16), deck,
                               fill=sc.get("stripe", FG), width=L(2))
            cab_window(cab0 + L(2), cab_t + L(2), cab1 - L(2),
                       cab_t + L(8))
            if sc.get("wide_nose"):
                cv.create_line(cab0 + L(1), nose_t, cab0 + L(4),
                               cab_t + L(2), fill="#bfd6e8", width=L(2))
            if sc.get("roof"):
                cv.create_rectangle(cab1, hood_t, x1 - L(2),
                                    hood_t + L(2), fill=sc["roof"],
                                    outline="")
            cv.create_rectangle(cab1, deck - L(3), x1 - L(2),
                                deck - L(1),
                                fill=sc.get("stripe", FG), outline="")
            for fi in range(sc.get("fans", 0)):
                fx = x1 - L(10) - L(fi * 11)
                cv.create_rectangle(fx - L(7), hood_t - L(2), fx,
                                    hood_t,
                                    fill=_mixc(sc["body"], "#000000",
                                               0.30), outline="")
            if sc.get("flare"):
                cv.create_polygon(x1 - L(24), hood_t + L(1), x1 - L(20),
                                  hood_t - L(2), x1 - L(4),
                                  hood_t - L(2), x1 - L(2),
                                  hood_t + L(1),
                                  fill=sc["body"], outline="")
            if sc.get("long45"):
                for gx in range(5):
                    cv.create_line(x1 - L(26) + L(gx * 5),
                                   hood_t + L(3),
                                   x1 - L(29) + L(gx * 5), deck - L(4),
                                   fill=_mixc(sc["body"], "#000000",
                                              0.35))
            if sc.get("flag"):
                # UP "Building America" flag, mid-hood
                f0, fy = cab1 + L(14), hood_t + L(3)
                fw, fh = L(19), L(10)
                cv.create_rectangle(f0, fy, f0 + fw, fy + fh,
                                    fill="#f2f4f6", outline="")
                for ry in (0, 2, 4, 6, 8):
                    cv.create_rectangle(f0, fy + L(ry), f0 + fw,
                                        fy + L(ry) + L(1),
                                        fill="#c0392e", outline="")
                cv.create_rectangle(f0, fy, f0 + L(8), fy + L(5),
                                    fill="#1b3d6d", outline="")
            if sc.get("tunnel"):
                for gx in range(3):
                    cv.create_rectangle(x1 - L(29) + L(gx * 9),
                                        deck - L(9),
                                        x1 - L(23) + L(gx * 9),
                                        deck - L(4), fill=FIELD,
                                        outline="")
            fuel_tank(x0, x1, tank_c)
            lw = ((x0 + L(12), x0 + L(22), x0 + L(32),
                   x1 - L(32), x1 - L(22), x1 - L(12))
                  if sc.get("axles", 4) == 6 else
                  (x0 + L(14), x0 + L(26), x1 - L(26), x1 - L(14)))
        underframe(x0, x1)
        wheels(lw)

        # ------------------------------------------------ three cars
        words = ("RUN8", "DLC", "MANAGER")
        cx = x1 + L(7)
        carw, gap = L(88), L(7)
        for wi, (ckind, cbody, csec, clet) in enumerate(cars):
            c0, c1 = cx, cx + carw
            cb = deck
            cmid = (c0 + c1) // 2
            cv.create_line(c0 - gap, cb - L(2), c0, cb - L(2), fill=DIM,
                           width=L(2))
            word = words[wi]
            label_bg = cbody
            if ckind == "box":
                ct = ROOF_BOX
                cv.create_rectangle(c0, ct, c1, cb, fill=cbody,
                                    outline="")
                cv.create_line(c0, ct + L(2), c1, ct + L(2), fill=csec)
                for rx in range(8, 88, 16):
                    cv.create_line(c0 + L(rx), ct + L(3), c0 + L(rx),
                                   cb - L(1), fill=csec)
                ty = (ct + cb) // 2
            elif ckind == "gon":
                ct = ROOF_GON
                cv.create_rectangle(c0, ct, c1, cb, fill=cbody,
                                    outline="")
                for rx in range(8, 88, 12):
                    cv.create_line(c0 + L(rx), ct + L(1), c0 + L(rx),
                                   cb - L(1), fill=csec)
                ty = (ct + cb) // 2
            elif ckind == "hop":
                ct = ROOF_HOP
                cv.create_polygon(c0, ct, c1, ct,
                                  c1, cb - L(3), c1 - L(13), cb,
                                  c0 + L(13), cb, c0, cb - L(3),
                                  fill=cbody, outline="")
                for rx in range(14, 78, 13):
                    cv.create_line(c0 + L(rx), ct + L(1), c0 + L(rx),
                                   cb - L(1), fill=csec)
                ty = (ct + cb) // 2
            elif ckind == "cov2":
                ct = ROOF_COV2
                cv.create_rectangle(c0, ct + L(3), c1, cb, fill=cbody,
                                    outline="")
                cv.create_polygon(c0, ct + L(3), c0 + L(10), ct,
                                  c1 - L(10), ct, c1, ct + L(3),
                                  fill=cbody, outline="")
                cv.create_line(cmid, ct + L(3), cmid, cb, fill=csec)
                ty = (ct + cb) // 2 + L(2)
            elif ckind == "rack":
                ct = ROOF_RACK
                cv.create_rectangle(c0, ct, c1, cb, fill=cbody,
                                    outline=csec)
                for ry in (3, 22):
                    for rx in range(6, 86, 8):
                        cv.create_rectangle(c0 + L(rx), ct + L(ry),
                                            c0 + L(rx + 3),
                                            ct + L(ry + 2), fill=csec,
                                            outline="")
                ty = (ct + cb) // 2
            elif ckind == "well":
                ct = ROOF_WELL
                cv.create_rectangle(c0 + L(2), cb - L(4), c1 - L(2), cb,
                                    fill="#3a3226", outline="")
                cv.create_rectangle(c0 + L(5), ct, c1 - L(5), cb - L(4),
                                    fill=csec, outline="")
                for rx in range(11, 79, 6):
                    cv.create_line(c0 + L(rx), ct, c0 + L(rx),
                                   cb - L(4),
                                   fill=_mixc(csec, "#000000", 0.25))
                label_bg = csec
                ty = (ct + cb - L(4)) // 2
            elif ckind == "reef":
                ct = ROOF_BOX
                cv.create_rectangle(c0, ct, c1, cb, fill=cbody,
                                    outline="")
                cv.create_rectangle(c0, ct, c1, ct + L(3), fill=csec,
                                    outline="")
                cv.create_rectangle(c0 + L(2), ct + L(4), c0 + L(6),
                                    cb - L(1), fill=csec, outline="")
                cv.create_rectangle(c1 - L(6), ct + L(4), c1 - L(2),
                                    cb - L(1), fill=csec, outline="")
                ty = (ct + cb) // 2
            elif ckind == "amfleet":
                ct = ROOF_AMF
                cv.create_rectangle(c0, ct + L(2), c1, cb, fill=cbody,
                                    outline="")
                cv.create_rectangle(c0 + L(4), ct, c1 - L(4), ct + L(3),
                                    fill=cbody, outline="")
                cv.create_rectangle(c0 + L(5), ct + L(3), c1 - L(5),
                                    ct + L(7), fill="#262a31",
                                    outline="")
                # the locomotive's phase stripes continue down the car,
                # at the exact same height as on the unit
                for bi, bc in enumerate(sc.get("bands",
                                               (csec, "#f2f4f7",
                                                "#3d6fd6"))):
                    by = ROOF_GENESIS + L(8 + bi * 3)
                    cv.create_rectangle(c0, by, c1, by + L(2),
                                        fill=bc, outline="")
                # a body-color patch breaks the stripes for the word
                fpre = fit_font(word, carw - L(20))
                half = fpre.measure(word) // 2 + L(5)
                cv.create_rectangle(cmid - half, ct + L(7), cmid + half,
                                    cb, fill=cbody, outline="")
                label_bg = cbody
                ty = (ct + L(7) + cb) // 2
            else:  # viewliner
                ct = ROOF_VIEW
                cv.create_rectangle(c0, ct, c1, cb, fill=cbody,
                                    outline="")
                for rx in range(8, 84, 12):
                    cv.create_oval(c0 + L(rx), ct + L(2),
                                   c0 + L(rx + 6), ct + L(7),
                                   fill="#262a31", outline="")
                cv.create_rectangle(c0, ct + L(9), c1, ct + L(12),
                                    fill=csec, outline="")
                label_bg = cbody
                clet = "#1b2b4a"
                ty = (ct + L(12) + cb) // 2
            f = fit_font(word, carw - L(12))
            cv.create_text(cmid, ty, text=word,
                           fill=readable(clet, label_bg), font=f)
            underframe(c0, c1)
            wheels((c0 + L(11), c0 + L(23), c1 - L(23), c1 - L(11)))
            cx = c1 + gap


    THEME_ICONS = {
        "Boxcar Slate":
            "AAABAAUAEBAAAAEAIACzAgAAVgAAABgYAAABACAAKAQAAAkDAAAgIAAAAQAgAMQFAAAxBwAAMDAAAAEAIAAoCAAA9QwAAEBAAAABACAAFgwAAB0VAACJUE5HDQoaCgAAAA1JSERSAAAAEAAAABAIBgAAAB/z/2EAAAJ6SURBVHjafZPPi1VlGMc/z/O+5/46452ZhBwdGzByo4QkVNRGpxFn2oSuwn/AnW3atSl3ItRCl/kPuG7hBMVEGRSzEERBJC+ppYmgd2bueM6555znaXFtxgvmAy8vfJ8ffPjCVwAB/NDxhfl8MJg3swbuwkvKRDyIVK00/eX6DyvfAyIA+z9476tnT/tfVnk+mpSX7oM7ALHZpDU99c2d31Y/l4OLC0f7d++vZP011yTWz4leWVZW0p7s6tTsnpNxuLa+WBWFSwh1keXRzbYIzAwAVd0iEBEarVZV5jn5xsZSxF1xFwR57d13CBMpbgbuaKsJIliWgwiiimU5g+s3BUfdXSMxUg1L2vvm2PPFZ9TDEgG03WJteQUrhuw8+THlYHNEkkQenP2a9Ru3kBiJIzJHGg0sLyj766gI6s7G1d+pn2VML81TPumjISBpG2k2tw0dc1gECQF5jqtpBxcB1ZEeAqK6tQyg4/badtMdKyu8qsa0sZkxAhU07SB5gThIEondHVhQJImIjn5NO0iM+NYBM2KSMHz4iM1fV0l274KJlOreAyYOvw0ORe8eOpFi/XWK2z2Ku/dpJAliNjogQRk+6dM7d4H3PzlDbUaMCdeWL+F1zaGl07hDjAmr311ERUg6bcyMWLu7huDTb8yCOQ/lD+q5GeI/j5mc2QXuPNIe9exI2zm3F1TJnvYNd+TAsfnFtb/+Xi6zzCQExwzt7sA2BtuZeEETwNxJWq3Q3T3zaXjc+/PO62+92bXh8MO6LNXd1TYzdVA3G70XNXeNSaKdqclvT129cl7+i/PBj46cKLJswc0a/j9xFhFXkaqRpj/f/PGny4D8C4P8Owmtz/aoAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAD70lEQVR42rWW329URRTHP2fu3Lvtdre/sBSKJBZo1CYiiIYaAlESE3/VRH7E8Df4QGJ4NBLefDEGfdInHzQxxIREgoEEAR+0iaYIBNBCg5aWQn9Q+mN3e3dn7owP2263tvKjiSe5yZ2c7znfM2fOmTMCwOGDiiNHHcCuA/s78nH8lDdW8F54FBHxogOCuvTQr98cu1btUwAFuC17ujvisXufxDO53daYOlYgKgji2vrsL6nmpkOXv//hdw4fVALItvf2bJgeGOyZHhltMXERwLMyER2FZFpXF5rWte3sPX7iggZ87vadz+ecl5QOIjyCPKZrDwgkxprc3dF0EARf7jp3vEu2H9jXcfda35WZ0bFQ6UDmgSuSOVtnE5dZ1axWb2zfoUtTMxsTa6NKWgS8TcD7Ze15AL9oXQG5JCEuFju0iPhqDz5J0E0NoNRSEplz7Zc/Int/CpEFeiXi9YKtYI2hdV83jd2v4WaLiFKLd2DMQqQiZSIRcA6pTTF9vofRr75dZFf5896jo5BM1zZUFGGSBFUTkXiPUwJKMfzpFwx+/BlJqUSQSWOSBCKN1xprLNmubQR1aXySVAj0kjyXDD5J8NaC9xWwB8y9CVxcrJyRtxasK+O9x5fMkrTpZbpy4ateAxKGSOKW6qrX/26+xytDj/eP14OK/1nUclE+aAcP1T+MQKJwATjXFwtotagEvXNVWI+Eesk5LOqDpGSY7esn1b6eVDqNhJraNS0k03kIA9Z/+AE4R9DcgDOGdFsrGAuFmLCpnlxPL0lhFgmCpQTee4IwZOzr75g8fY7OzjdZ07mTkf4LtD2znbMnPqoE4ozlle7DjA9co6ltI5NDg1y+cIzS+ERZX5UqXZ0/5xyUHPf6bvDcC1n+6j2NiOLOjV4abSNXL/+EFqF5Qzvjt/4kNz7E9Ogtmta2M9Z3g0y6PEYWEThBMRd95olVIJCNGxiWPpLNayn036Tx2VXE12dY396OiBCk09zdlGeqOEaqbQ253HXWrluH1KQQD/mJCbxz3osoHTXUDyilLEKgU1ElVYXfLqILBWpaW8ifOou7P4WurSmHNRuTP3mG1POdMDxC/tJVwkwdiCDlQhAVBBKF4d/ivZend3Sdnxwa3mWKRSNKhfPbM6VS+Y4KAlRYVV0ieGsx1oIIURRVV5YNwlA3PNn2R8eBvVsFYOvedzZPDd7uyY2Mpm3J+MoVXl1y/3V1V+m89xJoLdnWFp9tXb370snT52V++m95+/WXCvcnj8ZT09tdkqgVDn1qspmLUUP9oSunzvw4P/QXPVte3P/uy8V8fpM86pMFIHEQKKJUauCtg+//fOTVN+y8z38A0nrdZDJMDkYAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAIAAAACAIBgAAAHN6evQAAAWLSURBVHja7ZddiF1XFcd/a5+ve+98ZLwzk8QJbSoFGwNVU2mIbZWEqAmtELGOxqItVZDiS0D6PuYt6IPkRRA1IJUmMFI1pqWFaOKUVNRiCZaEBtN0JmMMae8kc+/cc8895+y9fDi305uZa2bMB/jgggObc/5n7f9e+7/XXkvotvFHPSZfsgC7Dx6I3DsX++r1OoMYbsbqOAYHBwkHB5Jj+w/EAEzsM+w/qIACSBdeAP3Unse2p0nydBq3trksG+r6djOmAMbzmkGl8npQLv3yjaMvHe2eTwAZHx83k5OT7v5dO3+Q1BvPNmtzZEmCOuWWTRUxBj8K6RuuUhka+sWH1q975o/3jKTsP6je+Pi4Nzk5ae/ftfOH8dzVZ+emL9o8y53c9KKXmBR+bJq5+Oo1a4x5ILf2o1cOHZmcmJgwAvDAnsc+07jy7lTtwnRmAt9Hb9fsS8mATbN0eONd4cDoyNffePGVIwYgacbPNGtzijFyxybvKML4vtd4r+bazfi7AGb3wQORTZJtWZKIETHcYRMRY9upyZLkEw/t/cqYn71+uj/P80F1CiIfrN4IqN4+HXQJWlWxeV5Jna36Xhiq6RyXbuW24wTP3Dgg1jkAPJFFsfXEWUsYRUi3PxEVxfn/6acPP/5FypvvQ/N8maJxDvE8JAoLvmlW4N4n0hU58X3a0xepvfAimuVFZLvsOgJihKQZM7L9YUa//QR2oYk6hxiDGIOztlhxEJDPN0hnZkGV4O4NhCNV1FpcJyrG81DncHnOwMMP4poxl371O8r9fah1vQkAOFWCtSO4VoIuxMQLC4SVCmEpIpmvIyL0rRul8fczzP7oJ1hn2fDUXkb37MYtxCRxjBhDeaCfNG6RtVr4UUSwdqRnpHtugeadUHVWLsaAkWLcCbP4PmG5hDhFwqAIfS98Z6y57UnA3Ch7rSbNqupSCfPf+Lzj534l+z+B/1ECtysFr8JnTwLi+4XCrUWdW+7LueK9U1xn3AuHc6i14BTxvdXlAWMMay5kfLbxOawmaFnJm01My+BUUZtTagyT3ruV5PtfBiAaGCZqrKHdnEdVi1xRU4JyP66cU0qGeO38q0yvRECdEpVKvH36OMFPv8f6+7aSthrcveULeEGIyzP8qMzZ3z+H5wcYPyj+szlZO17E2TzDCyPeOnkYP4y4OnuON6cOUyqXr0vDvTOhgPgeZ157gb+cPMy6zZuY/Tj86+eH0TRj/ZNfpX7+JNNTp+gPisuolWfLcU99jfr5E0xPnWIwjAiiqFOG3mgLVBcLUb8UEaQpa3Y8wvypP7Pm0w8Sja2n9soJhnc8QvnVPxFVygDk9Xw57uU/LOKCcmkxFcuSjOjbNBUHgirG9whKpWIfjUGA7Nx5Bnfv4PKh51FrWfvE4+Rn/0G5v4+oUlmsr1fCqXOICFm7Td5OQVVUMH51ZCSevfBOrKr4UUT1Ixuvr17encOfucSG7zyJ5hZv7hqtc28ztvlj1wdvBZyq4vkecxf/SZrUEN9PoiiqC8Dmndt/fXV6Zk+WtK0XBst0YdOMYHQY8Tyyy1fwAh/tcbmsBuey3IpnpHrPxr+dPTG11QcoR9HP0uHql2oXplXpcbsJtC9dRgET+Lg0KxLLUhIr4ESEPE1t9a4NYalUOiQiRWNy/LdH3xrbvOlez/e3tObrWXGPi4gI7z/G8zGeV5wSEcR88G01OBR1eZ4Nrh0N+4arJzcNDe87M/55EUCY2CfbzsxEzfn6oda1a3sb79XI22khxttUFfuBX7Rm1erxcl9l719/c2yuu0pYPKFbHt31jbTdfjqN40/aPO+/hcZ00TzPawWVyptBpfz86WMv/7gz12JzypIOWAEe+tY3x/J2e0jT7NYWHwYExiyceu7IzNJOHODf4gG5FxPVN/wAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAMAAAADAIBgAAAFcC+YcAAAfvSURBVHja7VpdbFxHFf7OzNx7989erx0nrTFt2tqEEDdVgNCWpGmTOHL4Ux7IAhJFIKGAeEAWEqrgoYQgVFAFD2kFEgiJP6Go+AVRVNWoIQngKhGlFEqNXDt2iUQSx7Ede31378/MHB52vTKOs966ZOuHjDTSaq6+e883c86cvwVuNPIfljjaL/D2D8rn8/KGD28o/MBzBgB6DvZ2Sil3gvk9zNwYQkKAgHGr9UuvDp4cAQAc7Rc4dpwBcC0CVJm258C+PdJ1HtNB+AhbmzZRBGZuzJYDkK4LkjJ2PPectvapfz7/wsCSx7wSAcLRfsKx47bnYO8TMPbrxdlZ+DOz0EFo2FpuqN4IQdJ1ZDKbRXP7BkDJE1cuXjoy+Y/X/KUkqgTy+bwcGBgw2/t6B4zWh6+OT9iwFLCUUhAR3UDZbt5ggJnZGGOVo7Dhrs1SpVN/C6ZnHhn70O6FRXUSS4Xv6dv/bWPM4cmR0SgOI1KOkkREiy9s6ARARKQcJdmyvDJ6PtIL/o5ka+6XOHbc4mg/lU+gYrDbDuzdRaA/TY2dNzqMJElBYKybUbG/+LatWxwhxZFXB0/+JJ/Py+qtQlI+XpydpahYwnoTvmwTBLZWzl+8xGz5G11dXd7AwIAVGHjO7Ojd28Gx3uPPzLJUSq434RdVWEgpSvMFZmPemem664GqDbDjPAhrkzoILRZ1fj0OAqzWVochM7AfAMoECFt1FDFby1i/4ldHHATEhK0AoCoGQsxMq3qXBqtMLYO2FXlV3UcX64Z6YuE6qxJZnQABsAwdxfA2tUM6DuxNJkFEYGsRX74CEIGUrEmiNgHLgCB0HPkMmj64s2IxjTgCQml4FJM/+gX0fKEmCVHr3g2CALm+fcgd6mus/muLpl070f7pj0NrvTYVYmYoKZHctgXWL4KthZASQcEHMyPZnAEzg4hQnCtASgkvk6rayb+/8xREMYDyXPi+j82PfRnOhhw4NiBJ0GGM0PeRbG4GSQIRIS6FiEolpLJZmGtz8LrvhptOgWMN3OB2X92IrS2DK5Otvc6Y2Rjwsg+Ek1OgBR9eMolwfh5szHVC2GW7y8ywxlRzAiz+rpU6/L909rolR0G6LuA4EK678g6utlaHT715ZskAMYOYgZt4c62HnPcWgVsEbhG4ReAWgfVMYK1OiAAmKocYa81S6/j2qrFQT3gfOmZ3ISoWIJWD+bkZsLVopjawZZAgzF27CqUU0twCrhTwXtcGrDWgNVhrvLfwPjR5nbA6AgmBOAxRmJtBi2gHCVmOfos+in4BObkRQir4/iQurEJe1UostDHwZy5DugmYwgxAQDKZKm9ORVC2FummLEQl0LO2HIB99PM/AGwMEhLGGHjpLEwcVKI4glQOmrKtICGq71HKQaYpC6tjOMkmBIVp2CAESbm2cNpzHJwdPI5N3e9Hyzu6wdaAhIQ1GnGpUNZBqQASYKuhvCSkkwBbg86eh8qqwwwQwURl4a3RiIMFEAFuMg2rY7A1cJJpeJlEVW2ChRmc+u03oY2BUyOhUTWDMUchmprGb773CXRvPwgnkYaOSki1bMK2A5+DcFz4s5NINrVCJrIYOX0CMxeG4SRSsEYvO1EJa2Iks+3XY5OZKlZ5KbDVGBn+PcLJKTie+xZSSgak68AUS/j76RMgIgRG411f+RLmskOY+Nb3UZq4AKcth81f6we2W5z/1a8Ra11WqSWp9ZvFeq4LtYrwdd5CAEmBZHMTXNdFe/c9aDm4F//54c8QT8+i+7uPw93QhjeefBqpni1o3X0/HKWQbG5CKpNGKpOuH7vrA1UsOQr1VAjr8wNcNrIoiqAyaQBAcWwCbQf3IXfgYWw8/DHEU9PQCz5UrgWx1oC1YMuVWSe2NVfF1lveVG/mTnYdB+HlKzB+Ea29e3Dx58/A+EVMP/8HNO+4FzKdQmlsHI5S//v9t4Kth0C1KkcENhbJbDOaN20s56fLUjwOQoSDZ7Dx0cMgpTB7egit+x7CbV94FOHZl5HRFpmtW1ZMFdeEZYaQEv7sNSxMXV0qz5LKnBATynVBQtBi9cFNJaFjfX2tN+Eh+ssrABjtH+nFxk8eAsII4dBLCE/+GSqdBglaWQXWgGVmKEchXFiorjmexxI0USUgovislVJLR0kTawSFAiZfH6tduhk7D/XsIKi5CSiWoK/NQzqqrrBhLVirK5tJJFQiQQI4Uz6GcvvS3tu3/1xhcmrn/OVJK5SSbO0q5k+ANtDGQAoBUqr+2Gkt2LJTtF4mTa133nEt9v07R148VxD54YvlXhPjyUz7BhKOYmYGSQESNSYI5Cg4CQ/CdUBULk7VxKwVK0W5Q8NsWzpuJyj59MiL5wr5fF7K4eFhzufz8tSzv3ttU/fd705kMtv96ZkIRHJd9DoqDVUTx3Hujk7Hy6RfCcCfnRmb4OHh4WpPgHC0nzoHh7y2bMtQWCzumBp/Q7O1Qkgp3ramBwPWWgazznV2OKm21qtBHN8/evLMeMWHWVrO9Y7dD+SymcxPrTaH5i5eQjBfsFZri2Ud8kbsO0lJXjolWzpuh0x4L+ti6VP/+uPQ6KLdrtR3qQp5X1/vFyHoqzqMunQYIg6ChjU4AMBJJKA8D8p1L5OgH/ujE0+MjY2FS4VficDSNe7q6vLSW+55kJgfNtZu4wakoILBJAhENE5EQyKMzvz1hVNzS0IfW9eLav3FpdGjIsuKlvhf8kI2kqmbC/sAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAQAAAAEAIBgAAAKppcd4AAAvdSURBVHja7ZtrjF3Vdcd/a+9z7jn33nna4/GT1BBjHBygibF5ObkxJKEylZrSDP0QtVFRK9pIFc2HfmnaTKat2qpqhZBSRBqpSlS1Qb5pkj4gUgshg92ASXk04MFgAgZC/Rg/5nVf55y9Vz/cO+MZz4yZgXEZ6CzpfjkzZ+29/nu99zqWhVL/3aZEbF/btV0YOqIsR2rtcdeuXTI0NLSgV2QhTBm4V4EpoTddf21+dZSPRvDLQu4uDJVa1b/85FNj05/39fXZcrns3jYA0xlc+ck9HwqMuU2FTwuyTYV2VBcG4sUlbe0hE+SQok/6JP3uoUcfOwi4uQ5wIQBIf3+/DAwM+G2l3deH+fiPgT2iBFmjQVKtkiUpIrIsNEBRjLVEhQJBHGOMwaPPuUby50OPPvatC2mDzAOKAmy/5RNfCXLhl7IkDSaGT1EbHc1cmomqyjI4+dkbF/HGWo3airatZ7XE7W24zD0wWq184fUDT5zt7+83AwMD/kIACMCHd9/YRT6+P8yFd4weP6FjJ4a9d86KMc1TF5YnKaCK9x4R8fnuLl39gU0Wr8+k9fqdQ4MHnj0fBDP9/VKpZAE0jr6ei6M7ho++noy8eQzAmsCeU3ldpr+mGmACixhjqqfP2OMvHkkVPmKj6Dvbdu5YPcDIjIM30x3e4OBgtv3m0pdzUe5XTr5yNKmeOZszYSBTQr9XqLVXEwZktXp44sjLqYhcajo7HmDgXk/f3im5LUB/f7+57777/FV7Pr7DxNE3x0+cZPzEcGDCQN5Tgs/lF6zBJalN6/Wss3fNllWbNp4afvDhg319fXZoaEgNwKRaaBD8tWsk4diJYUzw3hd+UhtMYKmNjtmJ4VM+CMM/2bL7hjXlctkDYlpx0m+7pXS5WHPTxPAp9c7ZZevo3iYIIiLjp06rsaY7zudvA7RUKllT+uGzBsCKuR0IamNjzhjD++L0p5uCMaS1Oo1KVY1IH2AGe4tqBnuLCmCsudnVG7g0E0R435GAqpqkUhHghituvK5I+SFnKD/kryztbhPlQ0m1iqoa3ofytxIladTqijXtYRheMRkGdcJloQqdWZLyvqWWSfskUYHAhUHnVB4g1irgRIT3PbVkNEgGECyqNJ5Kn2RZWomqLtx5N+uZGQAsrHug0KjWcM4tn2pQFSNCnM+DkUVFsGAxwqvz4D2rdn6EaPMl4BzvdsRQVcRastNnGf/xM2TVGjbKgdclBKAlvIlyrP/Cb9B23UdnGIzOYz86j22p6uwNiiBGLuS/5l9DFRGh+7U3OXbP16j97H+wuXBBmhAs1HEkScLGX/0l2j92HenJ01Mn3yyPBVSbgp2XfMx4LgLeY/IRJopQ51DAGIM6RzZWmzeJmZP/9LWdI9q8id7f/Bxv/Ok9CzaDYKGnH7e30XbdDtzYOGItGBCELM1wSUIYx5jQzthkUq1jjCHMR61TB1PMM/7sISr/fYi4swMB6hMVwrVr6L5lN5pk5xXpc/BpCZ8lGVmjQRjH2CjEjYyT33oZ+c2XMH7kFcI4ektTCBZoaEgYImEwCxyXJNTHx7FhiAnO7Vy90qhUsGFImI9azzwS5Rh/9nle2/c91m1YjxjDsWPH6Nm2lVW/sAdtpDOUfS4+k+SShMbEBEEuN1nYNjUiF87SlnfuBFVhLqYiTRWdO/OaHSlUMVGOfBxjiwXEGDra2rGF/Nz85+NzgbVF9SJEgaV13TjnUK8IincO9e9Oi93w/5xWAFgBYAWAFQBWAFgBYAWAFQBWAFgBYAWAFQAWVhIvBYlgJktcmdbZeRcoWMymd07cQOTb8d4BzR5erTLBxNgIXbKGMIrQVgdGVTl79jhhLkeH9KBe8T4jTw+PV/bzapqijQYYQ61eJ05SbhrZTVIbR8ROw3w2H2DOtb3LCJMi+8w9SwuAiOBrNdLaGFFbF+pSRCzqlVwU07VqDTYIpzY3+U5H12pEZOq5MZa0Ns7Wm25n3dZriaI8AElSJyp2kiV1RMystc/nA8y7ts9S3Ng41pglAkABa2hUaxx+9B/ZfedfUB09hfoMEYMxBmPsHC/R1AidNJ3mM5cldPReQvfGLS1NagLjXUaW1FoAzOwzB2Fuzo2dW1tRdRS71/Pi4ANUXn+TcIGt8YWZgFeiOGbowD6iYhdXfurz5AodqHpUm13dtF5tbsR7TBAiIrg0mWpZ2Vx+ysxVFe+yc+xdBiKEcbF16vPwy5IpDZnkp62WuM9SXhx8gP37BgiC4CL4AECs5eC/fpXnfryPsLcHvKdH1nKy/ga7dv8WH7zxM4gYGpVRfJYSd6xCxHDy5af5wb98BbFvrZar6V00vx7p5cTYq1Re/xlBECCBXcK2+HkUFQtkZ0aonhgmMIaTyfOsvuJyjn2yk+P2MY7//bc5/fAgPk3puObDbLzr18nd0Ev0zAZeffDf6YxisuYY23kBRgmM4UTy3OL4xXlG3GECa5tt8HMWeJHyAK9IYIkKeYJ8TDGK6bn9NsKe1Rz75j6OfuufaNu+jTW33syx/Y/z2l9+FW0kdP/yXnrWrMHmckSFPLl8POP3tvmFIVGxcO4mSC9WGJzl4xSfOcK2IvGWS0mPD3PmkcfYsOdjbP7S72GiHGHPKn56/zdYf+hF2nZcTW79WsYOHyGXj+e4GuOd8dOLlQfI/HnBpF+QIECzDJ+m5Hp7MFEOdY7chnUo4BtJ8/8mfYAIiF5cfrOD0uIBUO/f8oYlGRml/tobtF97DR1Xb+do+Z8Je1YRbVjH0b/6GzrW9VLY+kGSk6dpHDuBEcGn2bxZ5VLzkwtc3MwEQJHpwqoqUVuRqFicnL2dw4MIrlojfeon6M6fZ+Ndnycbn+Dl+78BQMe6Xn7ui79DsL6Xyvd/QGQshU0b5o/PS81PIEtS6iOjs1Pt6QMSNgwUIbWBbb0neO+IikU616/FZxeeA3BvHKP+H/uJPv1xLvuzP2D9C0fwjYT85ZcRblhL8pMXcE88TeemjZgFZGhLxU+MUB+foHZ2pHkDDZPa4D20RmT69tpXyg+NXXXrLYfDQuEmRDytm0ZFUd9MRvQCAIi1JIOPw/gE4fUfpX3H1WAt7vQIySMHaOw/iDQSCAL8Aq7AloqfwcwwDVXVXCFvvPcV59wLAEHpZEUGQXH+YBjHN1lrVb1HjKE2Mkparb31TWsLnOzlnxI+9DCyqguMoGMTpGdHCMIQjFl4NblU/ETQ1ihPSwafKxQscOia3vVnX+q/2wS9vb0KkKZpOSrkvxi1FU1tZBQJLFmSkjWSRZW57sxZspPDeFXClkd3b3f8bon4TQ5YBFFO47Y2nHPfLpfLbseOHWFQLpcdelTM9r1Pu03rXmrrWb21OjrmRTEisuihIwJLODlHMDm19U5q/XfCT87VVd55be/uNh7NXKPxPYCnfnG3MwB9d/y+GRoaSkizL0cd7VLo7vLeuWkfzywySfKtGaCl6J+8E34t4dV5wnzsOtatNT7NvvbCY/95pK+vzzJwr7cAQ0ND2tfXZ3/44Pef7/nAJdvbVnVfVRsdS32S2oUUMMuWpNk3EGPcmks3B2LMq2m1+tnTd342G7rv73RGLVAulz19e61Wqr+tzj+39vItYZCPU59mvCdnh1snLyJuzWWbrY2jqjYad7z4o4Pj0/PD6cer/Vfu1KEnnjzTqNc/p8KRtZdvCfNdnZnPnE5NcCzL78Vm7ktV8ZnTII6y3i2X2VyxMJ7V6nc9P3jgvyZVf/5Mf/IDip07VofdXf8QhMGtE6fPMD58KktrdVFV08wwlx8K2ox1PohyWuzuDjrW9YLyUlqp/NrQ/h89WSqVgsHBweytS50WCADbb/nE7wa58A9FpDepVGlUKiTVmmZpqstqVNZacoW8iQoForY2EBouc39bPXPmj1556tnRxXw4ee5vehRks1563bVri50de8WYzwA3Aj1mMYnN/xUQXsdUOCRev+MbjX97fvDA4fMPdDEAALM/Od2ya0dHnM9v1ThqI3PLRniv6lQYOvzI4Olzm99rKT/kWYKALKVSKaBvr30v+P9SqRTQf/eC4vf/AoaQWKoLf3K9AAAAAElFTkSuQmCC",
        "Amtrak Phase III":
            "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",
        "UP Armour Yellow":
            "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",
        "ATSF Warbonnet":
            "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",
        "BNSF Heritage II":
            "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",
        "SP Daylight":
            "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",
        "Conrail Blue":
            "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",
        "CSX Blue & Gold":
            "AAABAAUAEBAAAAEAIAC4AgAAVgAAABgYAAABACAAKwQAAA4DAAAgIAAAAQAgALEFAAA5BwAAMDAAAAEAIAAWCAAA6gwAAEBAAAABACAA9QsAAAAVAACJUE5HDQoaCgAAAA1JSERSAAAAEAAAABAIBgAAAB/z/2EAAAJ/SURBVHjafZNPa1xlFMZ/57zvvXOTzEySBlJCCbhQkBbSLkrVrFoV46ILFYoI+QpudCnin6XQfoG2H0Cwohux0FLQUkIXLVizqaSlYMG2mkmd5N479945x8VM0grRF87m4TzP+5x/AgjgR5ffOlXm26fMLMVd2OcZ4kGlySamfv5l7cplQATgpaXXPs/7vc+aqhyn7ssHHICYtMg6s+c27qx9LEdeXTm59ejBtaK/5RqT4f+wnzlpaploT+vM/KF3Y5U/XWmqgYuG4aAsopvtGsPMAFDVsQFHREhbWVNXJeVO/+2Iu4ILIjLz8kk06+JmCI4kE4DgdY4jiCpe5eT3bgju6u4a0UhTV2QLh5k9fZZhPUAEJJ2kuPMtVpd0jq/SFH+jGvCQ4t9/SP/+TSRE4siZI7GF1TnNTg9VATMG698xrHLaS+/RbP+FhgBpG+IE+Kih+qzBPqpdA6IR0QDtg2g6BaLILi4Kbns0/feUDBkrizuUPdyavfGBj8m+j4AoknYgpqNfYoJ0DxGzNhLSkaOQIK0OEpLdCoi4EWNC/ecG1cZVwvQitLp47z7ZC8sITv3kLtrq4sUmzaNfqf9YJ00TxG0kIBqotjd5+M1HfLmqeFGTJIEvfshwG/LJmRqKIUkS+PRSQFVIWpOYGXHo7hqCzx5cBDcu3D1KfmCJqa11puduAc7F347tYXMLt0GUot8zcOTwK2+uPH38+4/1oDDR4LihWRcr+3sb+TwmAmZO0spCd27h/fDk4b2N+cUXu9ZUy8OmVndXG+yog7rbOJ7HXGOS6GRn5vwHZ69/JbvnfOTE6+8MyuINN0v9P85ZRFxVmjSb+mn95tWvAfkHH0ExXmVzkbAAAAAASUVORK5CYIKJUE5HDQoaCgAAAA1JSERSAAAAGAAAABgIBgAAAOB3PfgAAAPySURBVHjatZZPbFRVFMZ/57773kxnWiothQiWEKDqwj8laoCQNpZgNMGVaJSFGxN3GmPC1hDcuGJBNJqwMS5kRzAYoy5MUYQajRYIaWEhUMBQCmXsv+mb9+67x8XMtEM6IjTxJC93ce/9zjnf+e45TwDo2284ecAD9O/e2zNXjjeoTwVV4X5MRMVYgqh4/dfvvxxpxBTAAL73+Vd64ulbB+O5mZ3OpUWWYcYEcUvritO51o59504eH6ZvvxFAntn1+sbpibGh6cmbXWkSAyjLM7FhRGvHmvLKrnV9vw8e+8MCOjv518c18MQENkJVEHkwaFUQIXNpOjs5XghMcLj/g8FtsvXFvT3jl0fOz5QmQhNYqR9cltXu+sz51vZOs7p70w6blKc2ZZmLFmgRQTPXlCXVek3/hR9jFw55nxFXKj1WRLQRQX2GLa4EMU2c1JGbl8jNlZAG70ZE7aLSBOdSOne8TaH3DXwyjxhzN0KWVNcgrDnT6qoeCVuYv/AtpcGDiCzeM4vBKzaMyG0aQGyOxDlMmMd5JcOAGEpf7+P2sXfJ0gom10aSOjSIULGkaUp+8wBBroj6bMGBXZJnVkG9Q7O0ymfmUEAF0tlJfBpDliJ4yBziHd5nqFegsgTONilV7asmpyJIjXsJQsRnaI1nFalWXKTG2NLqmwfUIfqAb9DwP5tpFuU9H5LqPTP8bwdBruGgQoMiEHOXBNVniw5VkQX5NimyiJClCcmNcxS7esgXYiQIya9ch4+nUGNZ9epnoBlSWIV3KYXObshSSOawxQ7iPwfJkjJigqUOVJXAhkz9dIjy8BHe33mNF56u8MsF6HvC8NKhJxcC8VnK8XdGOXsp4/Fuw7mxiA+PrieZuV3d99pEpqp478EnTF65yIqWiK+GPCLC0KhjW9cIR0/F2EDoWLuR85cdYxOGS+OeR9el3Lp6kdZiodqHGupkvYqhFn3rQ6tAhLb2mE9G+yl3biEeH2FFtB03dZjuDZMIQpAr8FHpU2ZuDBN2rKdwaZqH1x5BbB5BmZu+g3qvihgbFdrHjDEOkcBGuQWqymO/YTQk374Wd+YLfLmEjVpqXW0eN/w5ufUDmJmrxFd+JMy3gghSFYKYIJAoDK+IqspjvTtO/D1xvT9NKqkYE9bTS5Ok2qNsgAnCBnUJ6h1p6kCEKIoalOVdYEPbvvqR0Z7+vVsEYMvAnqembl0bmr1zs+DSRBdbuNxD40v3VFWCwEpbxxpt61iz8+zP35yQ+vTv7Xv5ufJM6VA8O7XV+8wsc+iTL7adiYrt+86f/u6H+tC/67fl2V2vba/Mz20W9P7nps/ABERRbmz3m++dOvDWgKtj/gNeEtSOt98a6AAAAABJRU5ErkJggolQTkcNChoKAAAADUlIRFIAAAAgAAAAIAgGAAAAc3p69AAABXhJREFUeNrtl12IXdUVx39r7/N170xuJneMNVNsNVJpB1KNDSHaRkKEGiqaPjgQS1/Svkh9KBbbB1+GgIVWH0pe6kNDoLQU4VZrJSqFodXGtKKxKgRFUNJ8ENIk83Un98752nv14dyZzCTXZGqS0ocu2HAv53/O/u//Xvu/1haWxlfGLB+2HMCOx/bGfvqfA+12m0bCZ4p2Co1Gg6jeSA/s29MFYOu44eAeBRRAluAF0K9t27ktz9Ldedrd4l0xtOTZZwkFMMZ2wqR+OIxrv3339ZdeWjqfADI2NmZarZbfcPf9T6ed9hOd2UmKPEW95+pDETEEUczA6mHqq4Z+veaGmx59ffqWnIN71I6NjdlWq+U23H3/M9321BNTp4+5siy8cK2i+pIrc99tTztjzF2lc7ef+dv+1vj4uBGAu7bt3Do3deavk6eOFiYIA1Sv3fzLuAiuyPPhdV+MVq1Z+8i7B19+zgCk851HO7OTihi5bpMDqGJsYOdmzvlsvvMDALPjsb2xy9MtRZ6KMWK4ziEixhWZKfL0jnt27BoJipOHB8uybFQJJxdWL2Yhia9NHqhfIoTiyrKel74Z2DBSIxfPpGRpF2suL4jrnRJr5LIn1TlHFMfIUoFFVER98Gkvrf3694k+vxF1xbIkqvh5xFgIatX/MkN9UZEQAb2wHrEhxbmPab+5r/qWmGXCLiMgxpB2O6y540FW3fsjfNpG1SFiEWPwzlUrDkNcd4Zy8hMERZrrCRs3os7he6oYa1Hv8a4g+dJ9aDbHmTf2UasPorj+BAC8KkFjHZp3IJ+je75NVBskjhPS9gxihMHmOtrHDnPuwJM471m7/XGGNn0HzedIOx3EGOqDDbL5DkXaJYgSbGOkr9J9t0B9CWJBDNIbiEGMQURQBLEhUVJD1CNBXEnfU0rMAt72flvwRV8C5nLutZJzrarL9vxKjrhCAv+9+D+B/1UCeh2m0pUTEBMAHvUOVd8n+atnqMf7Hq5P8yLqejgHJlyZDxhj+Gawn2c3PbvQNaFie4dIKR0ERmGTcvYhW9l24xnQp0EEVb3gFYvkPU8dTfj5lQio98RxQuuQY1Ut4hujSnteeGizIw6hKIVaDM+9YYkCCHtvl07oZPRwQlFCEgm/PxRQi4Qjx4W9B0KSJKwUuawTiiA24Fd/cvzixYLP3TrKU+0nmZn4Gepyhrb/mPTIC5x45xUGk0rW+ewCbuq1vUg2y9D2n5AeeZ4T77xCoxYRRlxSqPoQ0MW9DKKEMM+pffVhOh9NUBt9gHDNF2i/9wKrNzxA7f1XiZN6pUDZXsQN3L69h3t+ERfGtcreFUSWO2Lgily8IqAYGxBGSbWPxlT2fvLv1O/czezET+l6R+Pex+Hse9QGBolr9cUqfSWceo+IUOQZZZGBqqiKCZrDN3RPHj/aVVWCMKY5cuvy7iU7S3D+E+yOcXAlUXaW8vRbjKwfXa7dFXCqirUBU/86QT6TIjZI4zhuC8Do5vv+MH362M4iS50No0vywhU5YWMtYizF7GlsEKJ9istKcL4snBgrzZFb/vHh23/ZHADUknhfvnr425OnjlYpcnF1EyGbPoUqmCDE53nPWOQ/wokIZZG75k03R0mc7BeR6mIy8eofPxpZP3qbtcHG+fOzRa+mi4iwMIyxGGurUyLS6xPkkvFpOFTVu7JoNG+MBlYPv/blm4d++MHwmFTd5NZx2VL7IO7Mze6fn5vZNTdzjjLPeqZybbriIAyrq1mjOVGrD+x6+88vTi1tw2XBrDdu/dZ38yzbnafdO50rB6/iYroY1tr5MK4fCZP6794/eOCXS/ZFpU/LogD3PPi9kTLPhtTnV7d2ExEG5vyhl39z/OKbOMC/ARVpmu3A40ztAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAADAAAAAwCAYAAABXAvmHAAAH3UlEQVR42u1aXWxcRxX+zszcu3/XXq9jx05IUiI5FDduUhdFzU+TFJHg0BdUmhUgpCJ+ChJFAqGC4IFGEZAHnqrAAyCk8vNAkAV9iISaKvw0UVIqkRDa4jTEciKShji2E3vXu3t/ZubwsNcrx7XXG7fZ+iFHmoedvefO+WbOOXN+LrAQ9eYldh4QeP+J8vm8XPDPBYU/P2gAoG/7vjVSyC0AP8DMzQFEAgSMWKP/8carxy4AAHYeEDh5kAFwPQAUD9u3de8u6bjf0YH/GLPNmCgEMzdHfgKkckFCRo6beE1be/jN0y8NzpKR5wNA2HmAcPKg7du+7xCs+V65cAulwk3o0DdsLTdVb4QgqVyZasmiNdcJCPW7G9evPT068nppNogagHw+LwcHB82m7QODxuj941dHbBBUWAopiIhA1FzNZwYzs7HGKuWgY/V6qZKZf/qFiceGc5+YnlEnMVv4vm0DPzTG7B+9fCGMwoCUciRRLDlzcwcAIiKlHMnWyhtXLoban+5Ptbb/FicPWuw8QNUTiA124yN7dxDh5NiVYaOjQJKQBGYsF4rtL+pe3+sIIZ9+49Vjv8zn87LmVUjK75cLtyj0y1huwsc2AWYrC+PXmNk+19PTkxgcHLQC5wdN//Y9q1lHu0qFmyyVkstN+BkVFkKKynSB2Zq1XlfP1poNsHC2gW1Kh75F0631zvyrNdrqMGAGPgYAVQCgXh2FzNbycpZ/hqLQJwb1AoCKDYSYmRa9XZqsMvUM2sbyqoaPTkdNvYmFchcFsjgAIoAtdBgi0dYFKR3YuwyCiMBsEU1dB0AgqeqCqA+ALUACHR//LlIb9iwY+92NIwivncPkyz+ArhTqghD1/K7v+2jp/zQy/Z9rrv5bjeSGPcju+ga01ktTIWaGkhLuB/rBQRFsDYR0UCkVwMxIe1kwM4gIpeIkpJRIpltqdjL2h2cgTBnKSaBUKqHrU89DeSvBVoNIQEcB/HIR6ZY2EEkQEaKggsAvIdPaDi5PQHVthJvKgE204OkvbsQcXw3xYGvBbOc8osHi9gWCqVFQNI1EMoWgWABbA57jyayObv/NFtaY6jtJAGwWFU+9R0r7zhnpQMIFpAMh3fl3kMQi71rc5u5ihsUgro45SdR7Sssh570H4B6AewDuAbgHYFkDWOolRGCiOHxYahS7+NqLhhK/2vwtbL3fQhtAScLEZAWWLTpzHhgMAuHGRBFKKbRnU+B40RUvdIKNBowGW43jW/Zj9QrAMEESwQ8jjN8qo7ujBVISCITpcoDCtI+uzlZIAkYnGQ8eyS4NABFBG4OrEwJgi0gLgAiZdCLem6qghhm5bApEAoYZ1lTnTzx3A6FRkCKEMRFyHqCtiAM1hqMEOnJpUBwEGmY4jkIum0GkAaksxqYkbOSDhFxaOJ1wHXz9Nx3Y0fs/rGqzADSUM7Ng9TlJgFQC2gpIspCyGvv0rQOAcE7gZgEwDAuABJIugQEYw1ASkGrmOYa2Eo8ffgDavAWnTkKj6iXVJB2EhTE89GwOX9xdhJcCKgFjVTvwzOMGAOPmNKHdIygyOHJS4V+XBDJJQM+JhIVgRFqgO2fn8AJK8CxehjaEn51YiWDyLThu4t2klAzpuDBhGT89GoGI4EcGH3ziR/j5qd0Y/f1XUBm7BMdbga4nD4NMgLf/9BQirSGIaiZIsUr6UdgwbyIxBrWI8I15IWaQkEh5rXBdF51rNyC96UmMv3wI0fQE1jz1a8jcOoy++E0k1jyMtr59cJRCymtFOuMhnfHugHegxkvSQSMVwsbuAWawtQjDECrhVTOut8+h5SOfRfpDe5Hb+gVEhTHoSgEy04FIa4BtNXuLR2O8nTXeRsub6k58sus6CKauw/pFeA/lMf6X52H9aRTOHkGm51GIZAvCq2fgKDVn/XfD2wCAWlWOCGwNUl4Wre1dsNbcfgkRgSMfdOFFtO34Gkg6KL3+R7RufgJtH/02+PLf4CkNb33vvKWSpfEyhJAoFW9h+tbY7ArhrMociUvKcUEkiNlCCAk3mYbWVcO9jdwkzPBLcIMJ5B7+Etq2fRmkK+DLx2HfPAKVzIBiV/gOWgIvM0MpB0FlujbnuAmWhEs1AMKGf7dCaqkcaXQEv1zE6H//U/fo9JVhqLNHQclWICxDlychldNQDXUpvNboeDNJKDdJgvBK9Riq7Uv74LaB14o3R7cUJq5bIZWcWzqZt6JgNbQ2kFKAhGo8dloSLwFgm0h51L7qvsnIL9134dzposh3D8W9Jv6xl+skoRzm2HUSiYVHXDpxEkkI5aLaB6T6PEvlFTLu0LBt61xNEOonF86dLubzeSmHhoY4n8/Lvx47+u+utRs+nEx7m0pTEyGIJC2HXgERwAyjoyjXvc5JpLxzvsbnb14b5qGhoVpPgLDzAK0ZP5ZYkcueCvxy/9jVEc1shRBSvG9ND2ZYaxlgnVu5xklnV4z7QfTIxTN/HonvMDu3DMbr+h7NZVu9F6zRn5wauwa/VLDWaIs5HfK7ve/V6qKkRCoj2zpXQ7rJszoof+b8mRMXZ+x2vtpdTcjN2we+ChLP6ijo0WGAKPSb1uCousoklJuActzrROIXpdGLh4aHh4PZwi9UfJyZ456enkSm+/5tBN5trN3ITUhBBZirxksjRHRK2OCVM6eOT80KfWxDL6r3iUuzKZZlXkP8Pw5tQ5jkm4HaAAAAAElFTkSuQmCCiVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAALvElEQVR42u1baWxc13X+zl3eMgspitSQimQ6tC1bm7NUMW1JSceO66ao3SJ1OipQIwWaP2mLFEV/tAhgNzRbpG5/9E8DqGm6AC3aJMh4SVDLaVBUzlRJG1t2IjvyUJYlx6Zl09q4DWd5y72nP95wXzKUKZgSeIAByPdmzrvnu2e/50m0Sp8YEPk+yDdT/YRLZcZ6pOYa+/v7qVwut/QTaoUpjg0ygBmht++6y+/Muu54Y33IvckDqrW6PXPy+cm51wuFgiwWi+aKAZjLYPedv7RLCXE/g36ZCDsZlAVzayBeXeLmGmIivMKM520cPvXK8WefA2CW2sBWAKCBgQEaHBy0O/fl79Ku/+cA7iGwisMAYaOGOApBROtCAxiAEBKun4JyPAghYBk/NVHwWPn4s99YSRtoGVAYAPb03/uo0s7DcRSqqfGLqE9NxCaOiJlpHez84oUTWSEku6mMzGzqIi+VhTHxNycqtT8YPvmDsYGBATE4OGhXAoAAYO/HPr4Jyv+q1s6hicvv8uTl89ZaI4lEsutE69IHJibJsNaCiKyf7eDOnl4Jtj+Jwsbnyi+UTiwEQcz9fT6flwDA0vsHx/UOXXznjXD8wtsAIIVUsyrPvD4/zT1M1ipEbeKyfPfNVyMGPiq1++TOD93ROXh0/saLuQ6vVCrFe/o/+SXHcT9z4dzrYW1y1BFK0yy61wg11yqURhzU9fnh1yIi6hN++zdxbNBiV2FGbgkAAwMD4vDhw/b2/nv2Ce39S2X0PCqj55VQmq4pwZfyC0LCRKGMwkbcvjl3y+bc9ksXXzryXKFQkOVymSUAlOzdhOEScjfs+Lo10U2jI8NMQghcJ0RCIgrqpJRm108fyG7e+s8/OPqfVQAkmnHS7rzj3h0k5MGpsYtsrZHr1tFdoUkQEVXGL7EQssPz/fsBcD6flyIvvi8SL0cPAlD16qQRQuBaV/3FIVIgCuoI6lUWRAUAonQhx6J0IcdIEolPmrABE0e0DkP8WiAAZhZho0oA9t/2kQNpDBWNwFDR7v5YPkPgXWGjBmYW15X6z0+UKGjUGUJmtda3TYdBnmoYzaD2OA5x3VLTpG0cMgHKkG6fyQNISAZgCITrnxIZBSEGALWq0ng2rqxLK+F5GeHP/TItBKAlRwJmBI0ajDHrpxpkhiCC5/sArS6CqdUIz9YAbLFp593QXTsANu97UcjMSbY3dQG100cRBzVI7QJs1xCApvBCudj8K4/Cu/kX59VRzEsXiNMbsfAes11igQQSciX/tcIzLIgE0r/wEMaf/hPUL5+D1E5LmqBadRxhGCJ38Pfh3XofTOXdmZ1PymMBZpvY4IIUFGxh50rAFqR9CO2BjWk2MwTYGpjaxLKp7Dw+s2GtqfIW1hrorlvQdu8jCJ74QstmoFrdfS+VhXdzHlwfBwmV7BgBURTBRHU4ng8h9TwQgkYVUkho129eZ5BOofbmcwiGn4eX6QAR0KhOQrZtQ2bvA4CJFpnVYj6J8FEUIg5rcFwfUrmwtVE4PXvgb/kgKiOvQTvezzWFFjWAQVID0lmkGSYK0KhOQGkHQqh5thnUKpDKgXb95Jo1ENpD7Y0f4Z3SP6Jn6wdAJDAyMoKu3p3IfvhBcBwAJFfkM00mChDUJqEdd76jls4ibXzvThAMWqqvSAQisWz+TYsdAITy4XsepJsGkUBbNgPppJZV2yX5rPBsuipRYE3JwhgDZgsCYG3y9/tB103NvwHABgAbAGwAsAHABgAbAGwAsAHABgAbAGwAcPUAWLPjMgEhCARqlrTv38DJKsphwg8P/BrafMAwIAggECarDYyO19GzJQPP0eBmz8Bai3fOV+C6GrnNaTAYcczQkvHls4xHgihpfgiBer0Bz4R46c5+xFZAilmgl+KTrGbxs2MLKGHR+/hH1hYAIoKN6pioEtp8C2YBEIPBSHkKzpYMtBIziwOSPt+WzqThMX1dyGTHP3uPxcHdPlJeBQBQDxx0ZM42nzVfy5bik7Rnlnp2okmmMQnZ4um+akntSSJo1PC17yn8xW8HUMIitgJESUvaXdDMtc11eo4GYMFzrsWW0JsT6M0xgHBGu5J7CU/b1LBpcp3Fy7SMmWdbBowlKIpx5AWF6qXhpE3WQpOlNRNgC9f18HffNdic1fjC/RZa2jl+gcAgWAbYMpRMhIqNBTUlkbBJv46mR/YW2j1DCTvnvyX4WZrRkBl+YEhKDrSPvKDwu4fboFTjaviABPGBfwvx18/0QGdzAFvcve0Ujr7Zh6/81lk8uN8AEghiQhQDGS8R5eU3BO7/273L9v3nUn7b0Ar8EoAW8stvexVHX/8AqheHoVQdJNUatsUXkOunEU+NojZ2HkoKfOtsiM5exqP2WTz6fxJjxw5j4sQTsHGE9E370XXfF6E3bYfq+UsM//AbaE+7iWYsaHIy8xXy8/CdtwyUPAftequOVqvPA9iCpILrpaAcH2nfRXb/5yGz3RgtfQXnnv0q/Bv70bGvgAsvfw8Xn/wjcNRAat/voKtrC6Ry4HopOK4/73Pl/DRcPz17ErTKUH1lXWFmgABrYmg/A53bDTPxNiZfegrdH30APb/xNxDag8x2460jj6Hj7Z/A/+ABOO1bMfnOKTiuv9hBEb1HfnyV8oBlT4Bpxi+Q1GAbw5oIun1rcuxlDXTHjWAAHDUAIWd9wJKJz1rzay2BUytvtAXblZENq+MIL51Bqu8gMn134VzpnyCz3XA6bsTIt/8UbZ09cLfuRTz5LoKJEQgi2DjCMsPba86PVji4WQAA07xEgxmun4Hrp2HZYsnpERIwYQ3y7DOwfR9H131fhJ06j+EjjwEA2jp70P3rfwXRth1x+XG4SiCV2758fF5zfoQ4CtGYGl9CQ+YMSEilGaBISjWjjJYtXD+N9q6tsCZecUDaVN4CXv0OxG2fRvdD/4qOkZfBUQNOzx7I9hvAI8chzjyN9tx2tDJ/uVb8iAQatQrqlbHkBJoBSr5vLU+PyOwqyNd/XJy8ff+nTmk3dRAgi+YILU+bAdvpiZJl8wM79AScyZ8hvu0zSPUdAJMEVy9BnPkPmKEnQSYAhIK1tqV8Yy34CTHf/pmZHS8lrLVVY8wQAKh87gKVhsBg85x2vYNSSmZOBg7qlXFEyehcS44yfusM9IlnQP5mQAhwYxJRdQxK69WNrqwVv+bRPiUzggBgHS8lAbzy4ZtzY6fTA0LlcsmgZBRFRddL/bGbyoh6ZRwkFeIoRBwFq8kVYSqjiMcuwDJDKwkSCuaKx+/Whl/iBBlKu+ylMjDGPF4sFs2+ffu0KhaLBg8ziaf2/Njo7aczm7purVUmLDGL2QmMVcRYoaClnsnop2uFK28dvAd+zaGuJMcwnO3oEJYRmyj4NgC8mHrACAAonD4kyuVyCBt9yU23Uaqtw1prZhmsNkmangFaiwbKe+HXFJ6tgXZ909bVI2wc/f3Qi//zWqFQkDg2aCUAlMtlLhQK8vv/deRkV0/vnkz75tvrUxORjULZSgGzbokIbC2IhNmyrU8RiZ9FQe03L+/4XFz+7mGeVwsUi0WLXQXJYfX32Jqfdvfu0Mr1IxtHuCZnh2cdoNmy/SYpHa/GcXDo1RP/W5m1p/nFEA8c2s3lEz8aDYLGQwx6rbt3h/azm2JrYp6Z4FivL03NWRczw5qYlePFuRtukY6XrsRB/fMnj5demFb9+Qn4XJp+geJDd3TqdMe/K6U/NTVxGZWxi3EU1ImZRZJhrj8QOIl1VmmX020dqq2zBwCfjhrVz5ZfPPZ8Pp9XpVIpXlyBLAMCAOy5894/VMp5hIhyYaOKoF5F2KhxHEW8rkZlpYTjpYTrpeCmMgAoMCb+Wm1y9M9eL784sZoXJ2fvPczAl4n79tzZnc62/yoJ8WkABwB0rce3SpjtJINeIbZP2jh4+uTx0qmFG7oaAAAsfuX0lr39bZ7v38rSy8DG60Z4y2wYVD51/L8vz1zcVZAYKtplS8XVuJh8Pq+wq3AtxMRkrZ8YaKnb9f9Namg9zgwrYgAAAABJRU5ErkJggg==",
        "Rio Grande Gold":
            "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",
        "Chessie Yellow":
            "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",
        "BN Cascade Green":
            "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",
        "High Contrast":
            "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",
    }

    def write_theme_icon(dest, name=None):
        """The original icon, faithfully recolored per theme (pre-tinted
        at build time; the bright glyph carries the accent, the dark
        field keeps its depth)."""
        import base64 as _b64
        data = THEME_ICONS.get(name or THEME_NAME) \
            or THEME_ICONS["Boxcar Slate"]
        raw = _b64.b64decode(data)
        dest = Path(dest)
        try:
            if dest.exists() and dest.read_bytes() == raw:
                return dest            # unchanged: no write, no nudge
        except Exception:
            pass
        dest.write_bytes(raw)
        if sys.platform == "win32":
            try:
                # The icon file changed under the same path. Tell the
                # shell, touch the shortcuts so their icons re-extract,
                # and send per-shortcut change notices -- repaints the
                # Start Menu / Desktop without an app relaunch (Windows
                # may still lag a little).
                import ctypes, os
                sh = ctypes.windll.shell32
                sh.SHChangeNotify(0x08000000, 0, None, None)
                buf = ctypes.create_unicode_buffer(260)
                sh.SHGetFolderPathW(None, 0x0010, None, 0, buf)
                menu = (Path(os.environ.get("APPDATA", "")) /
                        "Microsoft" / "Windows" / "Start Menu" /
                        "Programs")
                for lnk in (Path(buf.value) / "Run8 DLC Manager.lnk",
                            menu / "Run8 DLC Manager.lnk"):
                    if lnk.is_file():
                        os.utime(lnk, None)
                        sh.SHChangeNotify(
                            0x00002000, 0x0005,   # UPDATEITEM, PATHW
                            ctypes.c_wchar_p(str(lnk)), None)
                sh.SHChangeNotify(0x08000000, 0x1000, None, None)
            except Exception:
                pass
        return dest

    class SetupPane(tk.Frame):
        """Settings content -- embedded as the Settings tab, or hosted
        in a Toplevel (SetupWizard) for the first-run walkthrough."""

        def __init__(self, master, parent, first_run=False, on_done=None):
            super().__init__(master, bg=PANEL, padx=20, pady=16)
            self.parent, self.first_run, self.on_done = parent, first_run, on_done

            cfg = load_json(DATA_DIR / "config.json", {})
            base = {**DEFAULT_CONFIG, **cfg}

            head = ("Step 1 of 2 -- folders and looks"
                    if first_run else "Settings")
            ttk.Label(self, text=head, style="Panel.TLabel",
                      font=parent.f_big).grid(row=0, column=0,
                                              columnspan=3, sticky="w")
            if first_run:
                ttk.Label(self, style="Dim.TLabel", wraplength=560,
                          justify="left",
                          text="The manager tracks which Run8 DLC you own "
                               "against the whole 3DTS store, keeps every "
                               "installer and transaction ID organized, "
                               "and handles reinstalls, quarantines, and "
                               "new purchases."
                          ).grid(row=1, column=0, columnspan=3,
                                 sticky="w", pady=(4, 4))

            self.vars = {}
            cur_theme = base.get("theme", "Boxcar Slate")
            self.v_theme = tk.StringVar(value=cur_theme)
            self._theme_was = cur_theme
            rowc = [2]

            def sect(title):
                ttk.Label(self, text=title, style="Panel.TLabel",
                          font=parent.f_big
                          ).grid(row=rowc[0], column=0, columnspan=3,
                                 sticky="w", pady=(14, 2))
                rowc[0] += 1

            # ---------------- folders ----------------
            sect("Folders")
            rows = (("run8_install", "Run8 install",
                     "where the sim lives -- auto-detected if standard"),
                    ("installers_dir", "Installers",
                     "your DLC installer EXEs (created if new)"),
                    ("backup_dir", "Backups",
                     "Back Up Now keeps one compressed "
                     "Run8DLC_Backup.zip here -- ideally another drive"))
            for key, label, hint in rows:
                r = rowc[0]
                ttk.Label(self, text=label, style="Panel.TLabel"
                          ).grid(row=r, column=0, sticky="w",
                                 pady=(4, 0), padx=(0, 10))
                v = tk.StringVar(value=base.get(key, ""))
                self.vars[key] = v
                ttk.Entry(self, textvariable=v, width=52
                          ).grid(row=r, column=1, sticky="we",
                                 pady=(4, 0))
                ttk.Button(self, text="Browse…", width=9,
                           command=lambda vv=v: self._browse(vv)
                           ).grid(row=r, column=2, sticky="w",
                                  padx=(8, 0), pady=(4, 0))
                ttk.Label(self, text=hint, style="Dim.TLabel"
                          ).grid(row=r + 1, column=1, columnspan=2,
                                 sticky="w")
                rowc[0] += 2

            # ---------------- appearance ----------------
            sect("Appearance")
            arow = ttk.Frame(self, style="Panel.TFrame")
            arow.grid(row=rowc[0], column=0, columnspan=3, sticky="w")
            rowc[0] += 1
            ttk.Label(arow, text="Theme", style="Panel.TLabel"
                      ).pack(side="left")
            self._tbtn = tk.Button(arow, relief="flat", bd=1,
                                    cursor="hand2",
                                    command=self._theme_menu)
            self._tbtn.pack(side="left", padx=(8, 0))
            self._paint_theme_btn()
            SCALE_CHOICES = {"Normal": 1.0, "Large (110%)": 1.1,
                             "Larger (125%)": 1.25, "Largest (150%)": 1.5}
            self._scale_choices = SCALE_CHOICES
            cur_f = float(base.get("ui_scale", 1.0))
            cur_lbl = next((k for k, v in SCALE_CHOICES.items()
                            if abs(v - cur_f) < 0.01), "Normal")
            self.v_scale = tk.StringVar(value=cur_lbl)
            self._scale_was = cur_lbl
            ttk.Label(arow, text="Size", style="Panel.TLabel"
                      ).pack(side="left", padx=(16, 0))
            ttk.Combobox(arow, textvariable=self.v_scale,
                         state="readonly", width=13,
                         values=list(SCALE_CHOICES)
                         ).pack(side="left", padx=(8, 0))
            self.v_cols = tk.StringVar(value=str(base.get("gal_cols", 3)))
            ttk.Label(arow, text="Gallery columns", style="Panel.TLabel"
                      ).pack(side="left", padx=(16, 0))
            ttk.Combobox(arow, textvariable=self.v_cols,
                         state="readonly", width=3, values=["2", "3", "4"]
                         ).pack(side="left", padx=(8, 0))
            if first_run:
                srow = ttk.Frame(self, style="Panel.TFrame")
                srow.grid(row=rowc[0], column=0, columnspan=3,
                          sticky="w", pady=(8, 0))
                rowc[0] += 1
                ttk.Label(srow, text="Which is easiest to read?",
                          style="Panel.TLabel").pack(side="left")
                for _lbl, _pts in (("Normal", 10), ("Larger (125%)", 12),
                                   ("Largest (150%)", 15)):
                    tk.Radiobutton(srow, text=_lbl.split(" ")[0],
                                   font=("Segoe UI", _pts),
                                   variable=self.v_scale, value=_lbl,
                                   bg=PANEL, fg=FG, selectcolor=FIELD,
                                   activebackground=PANEL,
                                   activeforeground=FG
                                   ).pack(side="left", padx=(12, 0))
            ttk.Label(self, text="theme and size apply on restart",
                      style="Dim.TLabel"
                      ).grid(row=rowc[0], column=0, columnspan=3,
                             sticky="w")
            rowc[0] += 1

            # ---------------- your data ----------------
            sect("Your data")
            drow = ttk.Frame(self, style="Panel.TFrame")
            drow.grid(row=rowc[0], column=0, columnspan=3, sticky="w")
            rowc[0] += 1
            ttk.Button(drow, text="Import purchase records…",
                       style="Accent.TButton",
                       command=lambda: (self._close(),
                                        parent.import_records_dialog())
                       ).pack(side="left")
            ttk.Button(drow, text="Back Up Now",
                       command=lambda: (self._close(),
                                        parent.backup_now())
                       ).pack(side="left", padx=(8, 0))
            ttk.Button(drow, text="Restore from backup…",
                       command=lambda: (self._close(),
                                        parent.restore_backup())
                       ).pack(side="left", padx=(8, 0))
            ttk.Label(self, text="import reads screenshots, emails "
                                 "(.eml) or documents; Back Up Now "
                                 "writes one compressed zip to your "
                                 "Backups folder",
                      style="Dim.TLabel"
                      ).grid(row=rowc[0], column=0, columnspan=3,
                             sticky="w")
            rowc[0] += 1

            self.v_store = tk.BooleanVar(value=first_run)
            if first_run:
                ttk.Checkbutton(
                    self, variable=self.v_store,
                    text="Download store data now (prices + product "
                         "images, recommended)"
                    ).grid(row=rowc[0], column=0, columnspan=3,
                           sticky="w", pady=(12, 0))
                rowc[0] += 1
            else:
                sect("Tools")
                t1 = ttk.Frame(self, style="Panel.TFrame")
                t1.grid(row=rowc[0], column=0, columnspan=3,
                        sticky="w", pady=(2, 0))
                rowc[0] += 1
                ttk.Button(t1, text="Refresh Store Prices",
                           command=parent.update_store).pack(side="left")
                ttk.Button(t1, text="HTML report",
                           command=parent.open_report
                           ).pack(side="left", padx=(8, 0))
                ttk.Button(t1, text="Transaction history",
                           command=parent.show_transactions
                           ).pack(side="left", padx=(8, 0))
                t2 = ttk.Frame(self, style="Panel.TFrame")
                t2.grid(row=rowc[0], column=0, columnspan=3,
                        sticky="w", pady=(6, 0))
                rowc[0] += 1
                ttk.Button(t2,
                           text="Update the Game (official updater)",
                           command=parent.run_updater).pack(side="left")
                ttk.Button(t2,
                           text="Create Desktop + Start Menu shortcuts",
                           command=self._shortcuts
                           ).pack(side="left", padx=(8, 0))
                t3 = ttk.Frame(self, style="Panel.TFrame")
                t3.grid(row=rowc[0], column=0, columnspan=3,
                        sticky="w", pady=(6, 0))
                rowc[0] += 1
                ttk.Button(t3,
                           text="Permanently delete disabled items…",
                           style="Danger.TButton",
                           command=parent.purge_quarantine
                           ).pack(side="left")
                ttk.Button(t3, text="Reset settings to defaults…",
                           command=self._reset_defaults
                           ).pack(side="left", padx=(8, 0))
                self.v_updchk = tk.BooleanVar(
                    value=bool(base.get("update_check", True)))
                tk.Checkbutton(self, variable=self.v_updchk,
                               text="Check the Run8 site for game "
                                    "updates and new DLC after each "
                                    "scan",
                               bg=PANEL, fg=FG, selectcolor=FIELD,
                               activebackground=PANEL,
                               activeforeground=FG
                               ).grid(row=rowc[0], column=0,
                                      columnspan=3, sticky="w",
                                      pady=(8, 0))
                rowc[0] += 1

            btns = ttk.Frame(self, style="Panel.TFrame")
            btns.grid(row=rowc[0], column=0, columnspan=3, sticky="e",
                      pady=(18, 0))
            ttk.Label(btns, text=f"Run8 DLC Manager v{VERSION}",
                      style="Dim.TLabel").pack(side="left", padx=(0, 24))
            ttk.Button(btns, text="Save & start" if first_run else "Save",
                       style="Accent.TButton", command=self._save
                       ).pack(side="right")
            ttk.Button(btns, text="Quit" if first_run else "Back",
                       command=(parent.destroy if first_run
                                else self._close)
                       ).pack(side="right", padx=(0, 8))
            self.columnconfigure(1, weight=1)

        def _reset_defaults(self):
            if DEMO_MODE:
                self.parent.log("settings reset (demo -- nothing "
                                "changed)", "warn")
                return
            if not messagebox.askyesno(
                    "Reset settings to defaults",
                    "This resets ALL settings -- folders, theme, text "
                    "size, window layout -- to fresh defaults and "
                    "reruns first-time setup on the next start."
                    "\n\nYour ledger, catalog and installers are NOT "
                    "touched.\n\nReset and restart?",
                    icon="warning", default="no", parent=self):
                return
            save_json(DATA_DIR / "config.json", dict(DEFAULT_CONFIG))
            restart_app()

        def _paint_theme_btn(self):
            t = self.v_theme.get()
            pal = PALETTES.get(t) or PALETTES["Boxcar Slate"]
            self._tbtn.configure(text="  " + t + "  ▾ ", bg=pal[0],
                                 fg=pal[6], activebackground=pal[2],
                                 activeforeground=pal[6],
                                 font=self.parent.f_big)

        def _theme_menu(self):
            old = getattr(self, "_tpop", None)
            if old is not None:
                old.destroy()
                self._tpop = None
                return
            pop = tk.Toplevel(self)
            self._tpop = pop
            pop.wm_overrideredirect(True)
            pop.configure(bg=FIELD, padx=2, pady=2)
            try:
                pop.attributes("-topmost", True)
            except Exception:
                pass

            def _shut(*_):
                if getattr(self, "_tpop", None) is not None:
                    self._tpop.destroy()
                    self._tpop = None

            for t in PALETTES:
                pal = PALETTES[t]
                row = tk.Frame(pop, bg=pal[0], cursor="hand2")
                row.pack(fill="x", pady=1)
                cv = tk.Canvas(row, highlightthickness=0)
                try:
                    draw_title_train(cv, t, self.parent.f_small, 0.72)
                except Exception:
                    cv.configure(width=200, height=30)
                cv.configure(bg=pal[0])
                cv.pack(side="left", padx=(4, 10), pady=3)
                tk.Label(row, text=t, bg=pal[0], fg=pal[6],
                         font=self.parent.f_big
                         ).pack(side="left", padx=(0, 16))

                def _pick(_e=None, t=t):
                    self.v_theme.set(t)
                    self._paint_theme_btn()
                    _shut()
                for w_ in (row, cv) + tuple(row.winfo_children()):
                    w_.bind("<Button-1>", _pick)
            pop.bind("<Escape>", _shut)
            pop.bind("<FocusOut>", _shut)
            pop.geometry("+%d+%d" % (
                self._tbtn.winfo_rootx(),
                self._tbtn.winfo_rooty()
                + self._tbtn.winfo_height() + 2))
            pop.focus_force()

        def _close(self):
            top = self.winfo_toplevel()
            if top is not self.parent:
                top.destroy()
            else:
                self.parent._set_mode(getattr(self.parent, "_last_view",
                                              "Gallery"))

        def _browse(self, var):
            d = filedialog.askdirectory(parent=self,
                                        initialdir=var.get() or str(Path.home()))
            if d:
                var.set(str(Path(d)))

        def _shortcuts(self):
            ok, msg = create_shortcuts()
            (messagebox.showinfo if ok else messagebox.showerror)(
                "Shortcuts", msg, parent=self)

        def _save(self):
            run8 = self.vars["run8_install"].get().strip()
            if not Path(run8).is_dir():
                if not messagebox.askyesno(
                        "Run8 folder", f"'{run8}' doesn't exist (yet?).\n\n"
                        "Save anyway?", parent=self):
                    return
            cfg = {**DEFAULT_CONFIG, **load_json(DATA_DIR / "config.json", {})}
            for k, v in self.vars.items():
                cfg[k] = v.get().strip()
            cfg["updater_exe"] = str(Path(cfg["run8_install"]) /
                                     "Run8_Updater.exe")
            cfg["_setup_done"] = True
            cfg["theme"] = self.v_theme.get()
            cfg["ui_scale"] = self._scale_choices[self.v_scale.get()]
            cfg["gal_cols"] = int(self.v_cols.get())
            if hasattr(self, "v_updchk"):
                cfg["update_check"] = bool(self.v_updchk.get())
            self.parent._gal_cols = cfg["gal_cols"]
            save_json(DATA_DIR / "config.json", cfg)
            theme_changed = (cfg["theme"] != self._theme_was
                             or self.v_scale.get() != self._scale_was)
            if cfg["theme"] != self._theme_was:
                try:
                    write_theme_icon(DATA_DIR / "run8dlc.ico",
                                     cfg["theme"])
                    self.parent.log("app icon recolored to match -- "
                                    "shortcuts follow (Windows may cache "
                                    "the old one briefly)", "accent")
                except Exception as e:
                    self.parent.log(f"icon update failed: {e!r}", "warn")
            if not self.first_run and theme_changed:
                if messagebox.askyesno(
                        "Appearance",
                        "Restart now to apply the new look?",
                        parent=self):
                    restart_app()
            for k in ("installers_dir", "backup_dir"):
                try:
                    Path(cfg[k]).mkdir(parents=True, exist_ok=True)
                except OSError:
                    pass
            want_store = self.first_run and self.v_store.get()
            self._close()
            if self.on_done:
                self.on_done()
            parent = self.parent
            if theme_changed and self.first_run:
                parent.log("color theme saved -- applies on next launch",
                           "accent")
            if self.first_run:
                RecordsStep(parent, want_store)
                return
            if want_store:
                parent.log("downloading store data...", "accent")
                parent._run_cli(["refresh", "--prices"],
                                done=lambda: parent._run_cli(
                                    ["media"], done=parent._media_done))
            else:
                parent.rescan()


    class SetupWizard(tk.Toplevel):
        """Toplevel host for the first-run walkthrough."""

        def __init__(self, parent, first_run=False, on_done=None):
            super().__init__(parent)
            self.title("Run8 DLC Manager -- Setup")
            self.configure(bg=PANEL)
            self.transient(parent)
            self.grab_set()
            self.resizable(False, False)
            if first_run:
                self.protocol("WM_DELETE_WINDOW", parent.destroy)
            SetupPane(self, parent, first_run=first_run,
                      on_done=on_done).pack(fill="both", expand=True)


    class RecordsStep(tk.Toplevel):
        """First-run step 2: how are your purchase records kept?"""

        HELP = {
            "screens": ("Put your receipt screenshots into the Receipts "
                        "folder you just chose (or pick them in a moment). "
                        "Windows' built-in OCR reads the transaction IDs "
                        "off them automatically."),
            "emails": ("In Gmail: open a receipt, menu > 'Download "
                       "message'. In Outlook: drag the mail into a folder. "
                       "That makes .eml files -- you'll pick them right "
                       "after setup and every transaction ID is read "
                       "automatically."),
            "docs": ("Text, CSV, Word or Excel all work. Each line just "
                     "needs the product and its transaction ID somewhere "
                     "near each other. You'll pick the file(s) right "
                     "after setup."),
            "skip": ("Totally fine -- everything works without them. "
                     "Ownership is detected from your installers and the "
                     "game itself; the manager, gallery, reinstall, "
                     "quarantine and store features are all fully "
                     "functional. You can import records any time later "
                     "from More > Import purchase records."),
        }

        def __init__(self, parent, want_store):
            super().__init__(parent)
            self.parent, self.want_store = parent, want_store
            self.title("Run8 DLC Manager -- Setup")
            self.configure(bg=PANEL, padx=20, pady=16)
            self.transient(parent)
            self.grab_set()
            self.resizable(False, False)
            ttk.Label(self, text="Step 2 of 2 -- your purchase records",
                      style="Panel.TLabel", font=parent.f_big
                      ).pack(anchor="w")
            ttk.Label(self, style="Dim.TLabel", wraplength=520,
                      justify="left",
                      text="The manager can keep every transaction ID "
                           "you've ever been issued next to its product. "
                           "How do you have your records?"
                      ).pack(anchor="w", pady=(4, 10))
            self.v_mode = tk.StringVar(value="screens")
            opts = (("screens", "Receipt screenshots (PNG / JPG)"),
                    ("emails", "Emails from the store"),
                    ("docs", "A document or spreadsheet"),
                    ("skip", "Skip purchase records -- just manage "
                             "my installs"))
            for val, label in opts:
                tk.Radiobutton(self, text=label, value=val,
                               variable=self.v_mode, bg=PANEL, fg=FG,
                               selectcolor=FIELD, activebackground=PANEL,
                               activeforeground=FG, anchor="w",
                               command=self._update_help
                               ).pack(fill="x", pady=1)
            self.help = ttk.Label(self, style="Dim.TLabel",
                                  wraplength=520, justify="left")
            self.help.pack(anchor="w", pady=(8, 0))
            self._update_help()
            btns = ttk.Frame(self, style="Panel.TFrame")
            btns.pack(fill="x", pady=(16, 0))
            ttk.Button(btns, text="Finish setup", style="Accent.TButton",
                       command=self._finish).pack(side="right")

        def _update_help(self):
            self.help.configure(text=self.HELP[self.v_mode.get()])

        def _finish(self):
            mode = self.v_mode.get()
            cfg = load_json(DATA_DIR / "config.json", {})
            cfg["records_mode"] = mode
            save_json(DATA_DIR / "config.json", cfg)
            parent = self.parent
            self.destroy()

            def after_records():
                if mode == "screens":
                    tx = Path(cfg.get("transactions_dir", ""))
                    imgs = []
                    try:
                        imgs = [f for f in tx.iterdir()
                                if f.suffix.lower() in
                                (".png", ".jpg", ".jpeg", ".bmp")]
                    except OSError:
                        imgs = []
                    if imgs:
                        parent.log(f"reading {len(imgs)} receipt "
                                   "screenshot(s)...", "accent")
                        parent._run_cli(["ocr-receipts"],
                                        done=parent.rescan)
                    else:
                        parent.log("drop receipt screenshots into your "
                                   "Receipts folder, then More > Import "
                                   "purchase records", "accent")
                elif mode in ("emails", "docs"):
                    if messagebox.askyesno(
                            "Import purchase records",
                            "Pick your saved receipt files now? (You can "
                            "always do it later from More > Import "
                            "purchase records.)", parent=parent):
                        parent.import_records_dialog()
                else:
                    parent.log("running without purchase records -- "
                               "everything else is fully functional",
                               "accent")

            if self.want_store:
                parent.log("downloading store data...", "accent")
                parent._run_cli(
                    ["refresh", "--prices"],
                    done=lambda: parent._run_cli(
                        ["media"],
                        done=lambda: (parent._media_done(),
                                      parent._when_idle(after_records))))
            else:
                parent.rescan()
                parent._when_idle(after_records)


def run_gui():
    if HAVE_TK:
        cfg = load_json(DATA_DIR / "config.json", {})
        apply_palette(cfg.get("theme", "Boxcar Slate"))
    if not HAVE_TK:
        out("This Python has no tkinter; the window can't open.")
        out("CLI still works: run with a command, e.g.  report")
        return 1
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        try:
            # Give the app its own taskbar identity: with an explicit
            # AppUserModelID, Windows takes the taskbar button's icon
            # from the live window (iconphoto), not from the launching
            # shortcut's cached .ico.
            import ctypes
            ctypes.windll.shell32.\
                SetCurrentProcessExplicitAppUserModelID("Run8.DLCManager")
        except Exception:
            pass
    Gui(None).mainloop()
    return 0


if __name__ == "__main__":
    if sys.argv[1:] == ["--demo"] and HAVE_TK:
        # Development/preview mode: the window shows a fabricated mix of
        # installed / owned / removed / not-owned so every UI state is
        # visible. Read-only -- actions that touch the library are
        # disabled, and nothing is written except normal window geometry.
        DEMO_MODE = True
        sys.exit(run_gui() or 0)
    if len(sys.argv) > 1:
        sys.exit(main() or 0)
    sys.exit(run_gui() if HAVE_TK else (main(["--help"]) or 0))
