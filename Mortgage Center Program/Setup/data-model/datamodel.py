"""Apex Project 03 data model generator.

The model lives in Apex-Data-Model.xlsx (sheets Objects and Fields). This script reads it and writes:
  python3 datamodel.py check            validate the workbook (run this after every edit)
  python3 datamodel.py package <dir>    deployable metadata in four folders, plus permission sets
  python3 datamodel.py guide|objects|events   the field tables used in the Project 03 guide
Needs Python 3.9+ and openpyxl (pip3 install openpyxl).
"""
from pathlib import Path
from xml.sax.saxutils import escape
import os, re, sys
NS = 'xmlns="http://soap.sforce.com/2006/04/metadata"'
MODEL = Path(os.environ.get("APEX_MODEL", Path(__file__).with_name("Apex-Data-Model.xlsx")))

def F(api, label, kind, desc="", history=False, **p):
    return dict(api=api, label=label, kind=kind, desc=desc, history=history, **p)
def obj(label, plural, fmt, sharing, history, fields, kind="custom"):
    return dict(label=label, plural=plural, fmt=fmt, sharing=sharing, history=history, fields=fields, kind=kind)

UI2KIND = {"Text": "text", "Text (Encrypted)": "encrypted", "Text Area (Long)": "longtext", "Currency": "currency",
           "Percent": "percent", "Number": "number", "Checkbox": "checkbox", "Date": "date", "Date/Time": "datetime",
           "Picklist": "picklist", "Auto Number": "autonumber", "Lookup": "lookup", "Master-Detail": "masterdetail",
           "Roll-Up Summary": "rollup", "Formula": "formula"}
STANDARD_OBJECTS = {"Account", "Contact", "Opportunity", "Case", "Lead", "User", "Campaign", "Contract"}

def _sheet_rows(ws):
    head = [c.value for c in ws[1]]
    out = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        d = {h: (v.strip() if isinstance(v, str) else v) for h, v in zip(head, row) if h}
        d = {k: (None if v == "" else v) for k, v in d.items()}
        if d.get("Object"): d["_row"] = i; out.append(d)
    return out

def _int(v): return None if v is None else int(v)
def _yes(v): return str(v or "").strip().lower() in ("yes", "y", "true", "1")

def load(path=MODEL):
    from openpyxl import load_workbook
    _ORPHANS.clear()
    wb = load_workbook(path)
    objects = _sheet_rows(wb["Objects"]); fields = _sheet_rows(wb["Fields"])
    standard, custom, events, rollups, formulas = {}, {}, {}, [], []
    for o in objects:
        k = o.get("Kind")
        if k == "Standard": standard[o["Object"]] = []
        elif k == "Custom": custom[o["Object"]] = obj(o["Label"], o["Plural label"], o["Record name format"], o["Sharing"], _yes(o.get("Track history")), [])
        elif k == "Platform Event": events[o["Object"]] = obj(o["Label"], o["Plural label"], None, None, False, [], kind="event")
    for r in fields:
        kind = UI2KIND.get(r.get("Type"))
        p = {}
        for col, key, conv in [("Length", "length", _int), ("Decimals", "decimals", _int), ("Default value", "default", str),
                               ("Display format", "fmt", str), ("Related to", "target", str), ("Child relationship name", "rel", str),
                               ("Roll-up child object", "child", str), ("Roll-up field", "field", str), ("Formula", "formula", str),
                               ("Formula return type", "rtype", str)]:
            if r.get(col) is not None: p[key] = conv(r[col])
        if r.get("Picklist values (; between)"): p["values"] = [v.strip() for v in str(r["Picklist values (; between)"]).split(";") if v.strip()]
        for col, key in [("Unique", "unique"), ("External ID", "extid"), ("Case-sensitive", "case")]:
            if _yes(r.get(col)): p[key] = True
        if r.get("Roll-up operation"): p["op"] = str(r["Roll-up operation"]).lower()
        if r.get("Roll-up filter field"): p["filt"] = (r["Roll-up filter field"], str(r.get("Roll-up filter values") or ""))
        f = F(r.get("Field name (API)"), r.get("Field label"), kind, r.get("Notes (description and help text)") or "", _yes(r.get("Track history")), **p)
        f["_row"], f["_obj"], f["_step"], f["_type"] = r["_row"], r["Object"], _int(r.get("Step")), r.get("Type")
        o, step = r["Object"], f["_step"]
        if step == 1 and o in standard: standard[o].append(f)
        elif step == 2 and o in custom: custom[o]["fields"].append(f)
        elif step == 2 and o in events: events[o]["fields"].append(f)
        elif step == 3: rollups.append(f)
        elif step == 4: formulas.append(f)
        else: f["_orphan"] = True; _ORPHANS.append(f)
    return standard, rollups, formulas, custom, events

_ORPHANS = []

def check(path=MODEL):
    """Return a list of human-readable problems; empty means the workbook is safe to generate from."""
    from openpyxl import load_workbook
    errs = []
    objects = _sheet_rows(load_workbook(path)["Objects"])
    known = {o["Object"]: o for o in objects}
    for o in objects:
        name, k = o["Object"], o.get("Kind")
        if k not in ("Standard", "Custom", "Platform Event"): errs.append(f"Objects row {o['_row']}: Kind must be Standard, Custom or Platform Event")
        if k == "Custom":
            if not name.endswith("__c"): errs.append(f"Objects row {o['_row']}: custom object {name} must end in __c")
            for col in ("Label", "Plural label", "Record name format", "Sharing"):
                if not o.get(col): errs.append(f"Objects row {o['_row']}: {name} needs {col}")
            if o.get("Record name format") and not re.search(r"\{0+\}", o["Record name format"]): errs.append(f"Objects row {o['_row']}: record name format needs a {{0000}} counter")
        if k == "Platform Event" and not name.endswith("__e"): errs.append(f"Objects row {o['_row']}: platform event {name} must end in __e")
    std, rollups, formulas, custom, events = load(path)
    all_fields = [(o, f) for o, fl in std.items() for f in fl] + [(o, f) for o, d in custom.items() for f in d["fields"]] \
               + [(o, f) for o, d in events.items() for f in d["fields"]] + [(f["_obj"], f) for f in rollups + formulas] + [(f["_obj"], f) for f in _ORPHANS]
    seen, rels = {}, {}
    for o, f in all_fields:
        where = f"Fields row {f['_row']} ({o}.{f['api']})"
        if f.get("_orphan"): errs.append(f"{where}: Step {f['_step']} doesn't fit object kind (1 = standard, 2 = custom or event, 3 = roll-up, 4 = formula), or the object is missing from the Objects sheet")
        if not f["api"] or not f["label"]: errs.append(f"{where}: needs Field label and Field name (API)"); continue
        if "__c" in f["label"]: errs.append(f"{where}: the label contains __c; type the readable name (e.g. {re.sub(r'(_c)?__c$', '', f['api']).replace('_', ' ')})")
        if not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*(_[A-Za-z0-9]+)*__c", f["api"]): errs.append(f"{where}: API name must be letters, digits and single underscores, ending in __c")
        if f["api"].endswith("_c__c"): errs.append(f"{where}: API name ends in _c__c; drop the extra _c")
        if (o, f["api"]) in seen: errs.append(f"{where}: duplicate of row {seen[(o, f['api'])]}")
        seen[(o, f["api"])] = f["_row"]
        k = f["kind"]
        if k is None: errs.append(f"{where}: Type '{f['_type']}' isn't one of {', '.join(UI2KIND)}"); continue
        need = {"text": ["length"], "encrypted": ["length"], "longtext": ["length"], "currency": ["length"], "percent": ["length", "decimals"],
                "number": ["length", "decimals"], "picklist": ["values"], "autonumber": ["fmt"], "lookup": ["target", "rel"],
                "masterdetail": ["target", "rel"], "rollup": ["op", "child"], "formula": ["formula", "rtype", "length", "decimals"]}.get(k, [])
        for n in need:
            if f.get(n) in (None, [], ""): errs.append(f"{where}: {f['_type']} needs {n.replace('fmt','Display format').replace('rel','Child relationship name').replace('target','Related to').replace('op','Roll-up operation').replace('child','Roll-up child object').replace('rtype','Formula return type')}")
        L, D = f.get("length") or 0, f.get("decimals") or 0
        if k == "text" and not 1 <= L <= 255: errs.append(f"{where}: Text length must be 1 to 255")
        if k == "encrypted" and not 1 <= L <= 175: errs.append(f"{where}: encrypted text length must be 1 to 175")
        if k == "longtext" and not 256 <= L <= 131072: errs.append(f"{where}: long text length must be 256 to 131,072")
        if k == "currency" and L + 2 > 18: errs.append(f"{where}: currency length + 2 decimals must be 18 or less")
        if k in ("percent", "number", "formula") and L + D > 18: errs.append(f"{where}: length + decimals must be 18 or less")
        if k == "picklist" and f.get("default") and f["default"] not in f.get("values", []): errs.append(f"{where}: default '{f['default']}' isn't one of the picklist values")
        if k == "picklist" and len(set(f.get("values", []))) != len(f.get("values", [])): errs.append(f"{where}: duplicate picklist values")
        if k == "autonumber" and f.get("fmt") and not re.search(r"\{0+\}", f["fmt"]): errs.append(f"{where}: display format needs a {{0000}} counter")
        if k in ("lookup", "masterdetail") and f.get("target") and f["target"] not in known and f["target"] not in STANDARD_OBJECTS:
            errs.append(f"{where}: Related to '{f['target']}' isn't on the Objects sheet")
        if k in ("lookup", "masterdetail") and f.get("rel"):
            key = (f["target"], f["rel"])
            if key in rels: errs.append(f"{where}: child relationship name '{f['rel']}' already used on {f['target']} by {rels[key]}")
            rels[key] = f"{o}.{f['api']}"
            if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", f["rel"]) or f["rel"].endswith("_"): errs.append(f"{where}: child relationship name must be letters, digits and underscores")
        if k == "masterdetail" and o not in custom: errs.append(f"{where}: a master-detail field can only be on a custom object")
        if o in events and k not in ("text", "number", "checkbox", "date", "datetime", "longtext"): errs.append(f"{where}: platform event fields can't be {f['_type']}")
        if k == "rollup":
            if f.get("op") not in ("count", "sum", "min", "max"): errs.append(f"{where}: Roll-up operation must be COUNT, SUM, MIN or MAX")
            if f.get("op") != "count" and not f.get("field"): errs.append(f"{where}: {str(f.get('op')).upper()} needs a Roll-up field")
            ch = custom.get(f.get("child"))
            md = [x for x in (ch["fields"] if ch else []) if x["kind"] == "masterdetail" and x.get("target") == o]
            if not md: errs.append(f"{where}: {f.get('child')} has no master-detail field to {o}, so it can't roll up")
            if ch and f.get("field") and not any(x["api"] == f["field"] for x in ch["fields"]): errs.append(f"{where}: {f['child']} has no field {f['field']}")
            if ch and f.get("filt") and not any(x["api"] == f["filt"][0] for x in ch["fields"]): errs.append(f"{where}: {f['child']} has no field {f['filt'][0]} to filter on")
    for o, d in custom.items():
        mds = [f for f in d["fields"] if f["kind"] == "masterdetail"]
        if len(mds) > 2: errs.append(f"{o}: an object can have at most two master-detail fields")
        if mds and d["sharing"] != "ControlledByParent": errs.append(f"{o}: has a master-detail field, so Sharing must be ControlledByParent")
        if not mds and d["sharing"] == "ControlledByParent": errs.append(f"{o}: Sharing ControlledByParent needs a master-detail field")
    return errs

if MODEL.exists():
    STANDARD, ROLLUPS, FORMULAS, OBJECTS, EVENTS = load()
else:
    STANDARD, ROLLUPS, FORMULAS, OBJECTS, EVENTS = {}, [], [], {}, {}

# ---------- XML ----------
def picklist_xml(f):
    v = "".join(f"<value><fullName>{escape(x)}</fullName><default>{'true' if x == f.get('default') else 'false'}</default><label>{escape(x)}</label></value>" for x in f["values"])
    return f"<type>Picklist</type><valueSet><restricted>true</restricted><valueSetDefinition><sorted>false</sorted>{v}</valueSetDefinition></valueSet>"

def field_body(f):
    k = f["kind"]
    if k == "text":
        return (f"<type>Text</type><length>{f['length']}</length><required>false</required>"
                f"<externalId>{'true' if f.get('extid') else 'false'}</externalId><unique>{'true' if f.get('unique') else 'false'}</unique>"
                + (f"<caseSensitive>{'true' if f.get('case') else 'false'}</caseSensitive>" if f.get("unique") else ""))
    if k == "encrypted": return f"<type>EncryptedText</type><length>{f['length']}</length><maskChar>asterisk</maskChar><maskType>lastFour</maskType><required>false</required>"
    if k == "longtext": return f"<type>LongTextArea</type><length>{f['length']}</length><visibleLines>5</visibleLines>"
    if k == "currency": return f"<type>Currency</type><precision>{f['length'] + 2}</precision><scale>2</scale><required>false</required>"
    if k == "percent": return f"<type>Percent</type><precision>{f['length'] + f['decimals']}</precision><scale>{f['decimals']}</scale><required>false</required>"
    if k == "number": return f"<type>Number</type><precision>{f['length'] + f['decimals']}</precision><scale>{f['decimals']}</scale><required>false</required>"
    if k == "checkbox": return "<type>Checkbox</type><defaultValue>false</defaultValue>"
    if k == "date": return "<type>Date</type><required>false</required>"
    if k == "datetime": return "<type>DateTime</type><required>false</required>"
    if k == "picklist": return picklist_xml(f)
    if k == "autonumber": return f"<type>AutoNumber</type><displayFormat>{f['fmt']}</displayFormat>"
    if k == "lookup": return f"<type>Lookup</type><referenceTo>{f['target']}</referenceTo><relationshipName>{f['rel']}</relationshipName><deleteConstraint>SetNull</deleteConstraint><required>false</required>"
    if k == "masterdetail": return f"<type>MasterDetail</type><referenceTo>{f['target']}</referenceTo><relationshipName>{f['rel']}</relationshipName><relationshipOrder>0</relationshipOrder><reparentableMasterDetail>false</reparentableMasterDetail><writeRequiresMasterRead>false</writeRequiresMasterRead>"
    if k == "rollup":
        filt = f"<summaryFilterItems><field>{f['child']}.{f['filt'][0]}</field><operation>equals</operation><value>{escape(f['filt'][1])}</value></summaryFilterItems>" if f.get("filt") else ""
        summed = f"<summarizedField>{f['child']}.{f['field']}</summarizedField>" if f.get("field") else ""
        return f"{filt}{summed}<summaryForeignKey>{f['child']}.{_md_field(f)}</summaryForeignKey><summaryOperation>{f['op']}</summaryOperation><type>Summary</type>"
    if k == "formula": return f"<formula>{escape(f['formula'])}</formula><formulaTreatBlanksAs>BlankAsZero</formulaTreatBlanksAs><precision>{f['length'] + f['decimals']}</precision><scale>{f['decimals']}</scale><type>{f['rtype']}</type>"
    raise ValueError(k)

def _md_field(f):
    """The child's master-detail field that points at the roll-up's object (Opportunity__c for the deal roll-ups)."""
    parent = f.get("_obj", "Opportunity")
    return next(x["api"] for x in OBJECTS[f["child"]]["fields"] if x["kind"] == "masterdetail" and x["target"] == parent)

def field_xml(f, event=False):
    desc = f"<description>{escape(f['desc'])}</description>" + ("" if event else f"<inlineHelpText>{escape(f['desc'])}</inlineHelpText>") if f["desc"] else ""
    hist = "" if event else f"<trackHistory>{'true' if f['history'] else 'false'}</trackHistory>"
    return f'<?xml version="1.0" encoding="UTF-8"?>\n<CustomField {NS}><fullName>{f["api"]}</fullName><label>{escape(f["label"])}</label>{desc}{hist}{field_body(f)}</CustomField>\n'

def editable(f): return f["kind"] not in ("autonumber", "rollup", "formula", "masterdetail")

def write_package(root):
    root = Path(root)
    def proj(folder):
        (root / folder).mkdir(parents=True, exist_ok=True)
        (root / folder / "sfdx-project.json").write_text('{\n  "packageDirectories": [{ "path": "force-app", "default": true }],\n  "sourceApiVersion": "66.0"\n}\n')
        return root / folder / "force-app/main/default"
    def ps(base, name, label, fields, objects=(), read_only=()):
        (base / "permissionsets").mkdir(parents=True, exist_ok=True)
        fp = "".join(f"<fieldPermissions><editable>{'true' if editable(f) else 'false'}</editable><field>{o}.{f['api']}</field><readable>true</readable></fieldPermissions>"
                     for o, f in fields if f["kind"] != "masterdetail")
        op = "".join(f"<objectPermissions><allowCreate>true</allowCreate><allowDelete>true</allowDelete><allowEdit>true</allowEdit><allowRead>true</allowRead>"
                     f"<modifyAllRecords>false</modifyAllRecords><object>{o}</object><viewAllRecords>false</viewAllRecords></objectPermissions>" for o in objects)
        # A child object's Read needs Read on its master-detail parent (standard objects such as Opportunity)
        op += "".join(f"<objectPermissions><allowCreate>false</allowCreate><allowDelete>false</allowDelete><allowEdit>false</allowEdit><allowRead>true</allowRead>"
                      f"<modifyAllRecords>false</modifyAllRecords><object>{o}</object><viewAllRecords>false</viewAllRecords></objectPermissions>" for o in read_only)
        # objectPermissions must stay sorted by object name
        op = "".join(sorted(re.findall(r"<objectPermissions>.*?</objectPermissions>", op), key=lambda s: re.search(r"<object>(.*?)</object>", s).group(1)))
        (base / "permissionsets" / f"{name}.permissionset-meta.xml").write_text(f'<?xml version="1.0" encoding="UTF-8"?>\n<PermissionSet {NS}>{fp}<label>{label}</label>{op}</PermissionSet>\n')
    def fields(base, o, fl, event=False):
        d = base / "objects" / o / "fields"; d.mkdir(parents=True, exist_ok=True)
        for f in fl: (d / f"{f['api']}.field-meta.xml").write_text(field_xml(f, event))
    b = proj("1-standard-object-fields")
    for o, fl in STANDARD.items(): fields(b, o, fl)
    ps(b, "Apex_Data_Model_1", "Apex Data Model 1", [(o, f) for o, fl in STANDARD.items() for f in fl])
    b = proj("2-custom-objects")
    for o, d in OBJECTS.items():
        (b / "objects" / o).mkdir(parents=True, exist_ok=True)
        (b / "objects" / o / f"{o}.object-meta.xml").write_text(
            f'<?xml version="1.0" encoding="UTF-8"?>\n<CustomObject {NS}><deploymentStatus>Deployed</deploymentStatus><enableActivities>false</enableActivities>'
            f'<enableHistory>{"true" if d["history"] else "false"}</enableHistory><enableReports>true</enableReports><enableSearch>true</enableSearch>'
            f'<externalSharingModel>{d["sharing"]}</externalSharingModel><label>{d["label"]}</label><nameField><displayFormat>{d["fmt"]}</displayFormat>'
            f'<label>{d["label"]} Number</label><trackHistory>false</trackHistory><type>AutoNumber</type></nameField><pluralLabel>{d["plural"]}</pluralLabel>'
            f'<sharingModel>{d["sharing"]}</sharingModel></CustomObject>\n')
        fields(b, o, d["fields"])
    for o, d in EVENTS.items():
        (b / "objects" / o).mkdir(parents=True, exist_ok=True)
        (b / "objects" / o / f"{o}.object-meta.xml").write_text(
            f'<?xml version="1.0" encoding="UTF-8"?>\n<CustomObject {NS}><deploymentStatus>Deployed</deploymentStatus><eventType>HighVolume</eventType>'
            f'<label>{d["label"]}</label><pluralLabel>{d["plural"]}</pluralLabel><publishBehavior>PublishAfterCommit</publishBehavior></CustomObject>\n')
        fields(b, o, d["fields"], event=True)
    ps(b, "Apex_Data_Model_2", "Apex Data Model 2", [(o, f) for o, d in OBJECTS.items() for f in d["fields"]], OBJECTS.keys(), read_only=("Opportunity",))
    b = proj("3-rollups"); fields(b, "Opportunity", ROLLUPS); ps(b, "Apex_Data_Model_3", "Apex Data Model 3", [("Opportunity", f) for f in ROLLUPS])
    b = proj("4-ltv-formula"); fields(b, "Opportunity", FORMULAS); ps(b, "Apex_Data_Model_4", "Apex Data Model 4", [("Opportunity", f) for f in FORMULAS])

# ---------- guide tables (UI terms) ----------
def ui_type(f):
    k = f["kind"]
    return {
     "text": lambda: f"Text, length {f['length']}" + (", **Unique** + **External ID**" if f.get("extid") else (", **Unique**" + (" (case-sensitive)" if f.get("case") else "") if f.get("unique") else "")),
     "encrypted": lambda: f"Text (Encrypted), length {f['length']}, mask *Last Four Characters Clear*",
     "longtext": lambda: f"Text Area (Long), length {f['length']:,}",
     "currency": lambda: f"Currency, length {f['length']}, 2 decimal places",
     "percent": lambda: f"Percent, length {f['length']}, {f['decimals']} decimal places",
     "number": lambda: f"Number, length {f['length']}, {f['decimals']} decimal places",
     "checkbox": lambda: "Checkbox, default unchecked",
     "date": lambda: "Date", "datetime": lambda: "Date/Time",
     "picklist": lambda: "Picklist (restricted): " + ", ".join(f"{v} (default)" if v == f.get("default") else v for v in f["values"]),
     "autonumber": lambda: f"Auto Number, display format `{f['fmt']}`, starting number 1",
     "lookup": lambda: f"Lookup Relationship → {f['target']}",
     "masterdetail": lambda: f"**Master-Detail Relationship** → {f['target']}",
     "rollup": lambda: f"Roll-Up Summary: **{f['op'].upper()}** of {f['child'].replace('__c','')}" + (f" → {f['field']}" if f.get("field") else "") + (f", filter {f['filt'][0]} = {f['filt'][1]}" if f.get("filt") else ""),
     "formula": lambda: f"Formula (return type {f['rtype']}, {f['decimals']} decimal places): `{f['formula']}`",
    }[k]()

def table(rows, history_col=False):
    out = ["| Field label | Field name (API) | Type and settings | Notes |", "|---|---|---|---|"]
    for f in rows:
        notes = f["desc"] + (" · Track history" if f["history"] else "")
        out.append(f"| {f['label']} | `{f['api']}` | {ui_type(f)} | {notes.strip(' ·')} |")
    return "\n".join(out)

def guide_markdown():
    parts = []
    parts.append("**Account (Household) and Contact (borrower) fields.** Object Manager → Account (or Contact) → Fields & Relationships → New. Type the field name without `__c`; Salesforce adds it.\n")
    parts.append("Account:\n\n" + table(STANDARD["Account"]) + "\n\nContact:\n\n" + table(STANDARD["Contact"]))
    parts.append("The SIN field exists to practise protecting a restricted field: a SIN isn't required to get a mortgage, and the better design is not to store it in the CRM at all (data minimization, Module 14 GV2).")
    parts.append("**Opportunity fields (the deal).** Create these now:\n\n" + table(STANDARD["Opportunity"]))
    parts.append("Create these three roll-ups **after** `Property__c` and `Condition__c` exist (Object Manager → Opportunity → New → Roll-Up Summary), then the formula:\n\n" + table(ROLLUPS + FORMULAS))
    return "\n\n".join(parts)

def objects_markdown():
    parts = []
    order = ["Approval_Request__c", "Property__c", "Applicant_Income__c", "Liability__c", "Condition__c", "Loan__c", "Loan_Party__c"]
    for o in order:
        d = OBJECTS[o]
        sharing = "Controlled by parent (set by the master-detail field)" if d["sharing"] == "ControlledByParent" else "Private (set in Sharing Settings, step 0.4)"
        rl = next((f for f in d["fields"] if f["kind"] == "masterdetail"), None)
        head = (f"**{d['label']}** (`{o}`): label `{d['label']}`, plural `{d['plural']}`, record name `{d['label']} Number` as Auto Number `{d['fmt']}`, "
                f"Allow Reports{', Track Field History' if d['history'] else ''}. Sharing: {sharing}."
                + (f" Related list label on the parent page: `{d['plural']}`." if rl else ""))
        parts.append(head + "\n\n" + table(d["fields"]))
    return "\n\n".join(parts)

def events_markdown():
    parts = []
    for o, d in EVENTS.items():
        parts.append(f"**{d['label']}** (`{o}`): Publish Behavior **Publish After Commit**.\n\n" + table(d["fields"]))
    return "\n\n".join(parts)

def main(argv):
    if len(argv) < 2 or argv[1] not in ("check", "package", "guide", "objects", "events"):
        print(__doc__); return 2
    if not MODEL.exists():
        print(f"Can't find {MODEL}. Put Apex-Data-Model.xlsx next to this script or set APEX_MODEL."); return 2
    if argv[1] in ("check", "package"):
        errs = check()
        if errs:
            print(f"{len(errs)} problem(s) in {MODEL.name}:"); print("\n".join("  - " + e for e in errs)); return 1
        if argv[1] == "check":
            n = sum(len(v) for v in STANDARD.values()) + sum(len(d["fields"]) for d in list(OBJECTS.values()) + list(EVENTS.values())) + len(ROLLUPS) + len(FORMULAS)
            print(f"OK: {len(OBJECTS)} custom objects, {len(EVENTS)} platform events, {n} fields"); return 0
        write_package(argv[2]); print(f"Wrote metadata to {argv[2]}"); return 0
    print({"guide": guide_markdown, "objects": objects_markdown, "events": events_markdown}[argv[1]]()); return 0

if __name__ == "__main__":
    sys.exit(main(sys.argv))
