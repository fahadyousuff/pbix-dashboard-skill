"""
Helpers to generate PBIR visual.json files (schema 2.12.0).

Usage:
    from build_visuals import mk_card, mk_bar, mk_line, mk_donut, write_page

    visuals = [
        mk_card("Sales",  x=20,   y=24, w=440, h=160, z=0, table="financials"),
        mk_card("Profit", x=480,  y=24, w=440, h=160, z=1, table="financials"),
        mk_bar(cat="Segment", measure="Sales", x=20, y=210, w=940, h=400, z=2, table="financials"),
        mk_line(cat="Month Number", measure="Sales", series="Year",
                x=980, y=210, w=920, h=400, z=3, table="financials"),
        mk_donut(cat="Discount Band", measure="Sales",
                 x=20, y=630, w=620, h=430, z=4, table="financials"),
    ]
    write_page(root="pbix_src", page_id=None, display_name="Dashboard", visuals=visuals)
"""
import json, os, random


SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.12.0/schema.json"
PAGE_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
PAGES_SCHEMA = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json"


def new_id():
    return "".join(random.choices("0123456789abcdef", k=20))


def _col(table, prop):
    return {"Column": {"Expression": {"SourceRef": {"Entity": table}}, "Property": prop}}


def _agg(table, prop, fn=0):
    return {"Aggregation": {"Expression": _col(table, prop), "Function": fn}}


def _col_proj(table, prop):
    return {"field": _col(table, prop), "queryRef": f"{table}.{prop}",
            "nativeQueryRef": prop, "active": True}


def _agg_proj(table, prop, fn=0):
    fn_name = {0: "Sum", 1: "Avg", 2: "Count", 3: "Min", 4: "Max"}.get(fn, "Sum")
    return {"field": _agg(table, prop, fn), "queryRef": f"{fn_name}({table}.{prop})",
            "nativeQueryRef": f"{fn_name} of {prop}", "active": True}


def _sort_desc(field):
    return {"sort": [{"field": field, "direction": "Descending"}], "isDefaultSort": True}


def _sort_asc(field):
    return {"sort": [{"field": field, "direction": "Ascending"}], "isDefaultSort": True}


def _wrap(vid, x, y, w, h, z, visual_type, query_state, sort=None):
    vis = {
        "$schema": SCHEMA,
        "name": vid,
        "position": {"x": float(x), "y": float(y), "z": z,
                     "height": float(h), "width": float(w), "tabOrder": z},
        "visual": {
            "visualType": visual_type,
            "query": {"queryState": query_state},
            "drillFilterOtherVisuals": True,
        },
    }
    if sort:
        vis["visual"]["query"]["sortDefinition"] = sort
    return vis


def mk_card(measure, x, y, w, h, z, table, fn=0):
    return _wrap(new_id(), x, y, w, h, z, "cardVisual",
                 {"Data": {"projections": [_agg_proj(table, measure, fn)]}},
                 _sort_desc(_agg(table, measure, fn)))


def mk_bar(cat, measure, x, y, w, h, z, table, fn=0):
    return _wrap(new_id(), x, y, w, h, z, "clusteredBarChart",
                 {"Category": {"projections": [_col_proj(table, cat)]},
                  "Y":        {"projections": [_agg_proj(table, measure, fn)]}},
                 _sort_desc(_agg(table, measure, fn)))


def mk_column(cat, measure, x, y, w, h, z, table, fn=0):
    return _wrap(new_id(), x, y, w, h, z, "clusteredColumnChart",
                 {"Category": {"projections": [_col_proj(table, cat)]},
                  "Y":        {"projections": [_agg_proj(table, measure, fn)]}},
                 _sort_desc(_agg(table, measure, fn)))


def mk_line(cat, measure, x, y, w, h, z, table, fn=0, series=None):
    qs = {"Category": {"projections": [_col_proj(table, cat)]},
          "Y":        {"projections": [_agg_proj(table, measure, fn)]}}
    if series:
        qs["Series"] = {"projections": [_col_proj(table, series)]}
    return _wrap(new_id(), x, y, w, h, z, "lineChart", qs,
                 _sort_asc(_col(table, cat)))


def mk_donut(cat, measure, x, y, w, h, z, table, fn=0):
    return _wrap(new_id(), x, y, w, h, z, "donutChart",
                 {"Category": {"projections": [_col_proj(table, cat)]},
                  "Y":        {"projections": [_agg_proj(table, measure, fn)]}},
                 _sort_desc(_agg(table, measure, fn)))


def write_page(root, display_name, visuals, page_id=None, make_active=True):
    """Add a new page with the given visuals to an extracted PBIX source tree."""
    if page_id is None:
        page_id = new_id()

    pages_dir = os.path.join(root, "Report", "definition", "pages")
    os.makedirs(pages_dir, exist_ok=True)

    page_dir = os.path.join(pages_dir, page_id)
    os.makedirs(page_dir, exist_ok=True)

    with open(os.path.join(page_dir, "page.json"), "w", encoding="utf-8") as f:
        json.dump({
            "$schema": PAGE_SCHEMA, "name": page_id, "displayName": display_name,
            "displayOption": "FitToPage", "height": 1080, "width": 1920,
        }, f)

    vis_dir = os.path.join(page_dir, "visuals")
    os.makedirs(vis_dir, exist_ok=True)
    for v in visuals:
        d = os.path.join(vis_dir, v["name"])
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "visual.json"), "w", encoding="utf-8") as f:
            json.dump(v, f)

    # Register page in pages.json
    pages_json = os.path.join(pages_dir, "pages.json")
    if os.path.exists(pages_json):
        with open(pages_json, "r", encoding="utf-8") as f:
            meta = json.load(f)
    else:
        meta = {"$schema": PAGES_SCHEMA, "pageOrder": [], "activePageName": page_id}
    if page_id not in meta["pageOrder"]:
        meta["pageOrder"].append(page_id)
    if make_active:
        meta["activePageName"] = page_id
    with open(pages_json, "w", encoding="utf-8") as f:
        json.dump(meta, f)

    return page_id
