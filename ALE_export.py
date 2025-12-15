#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import shutil

resolve = bmd.scriptapp("Resolve")
fusion = resolve.Fusion()
ui = fusion.UIManager
dispatcher = bmd.UIDispatcher(ui)

DEFAULT_SETTINGS = {
    "export_to": "",
    "create_export_folder_timeline_name": True,

    "export_format": "ale",  # ale | csv_edit_index | csv_metadata | ale_plus_csv_metadata

    "tracks_mode": "video_only",  # video_only | video_audio

    "merge_slate": False,
    "name_mode": "clipname",  # clipname | slate

    "slate_format": "scene_shot_take",  # scene_shot_take | scene_take_slash | scene_take_dash
    "slate_add_selected": True,
    "slate_add_camera": True,

    "tape_mode": "as_set_in_project",  # as_set_in_project | reel | unc_clipname_no_ext | empty
}

def _settings_path():
    script_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    folder = os.path.join(script_path, "ALE_export_settings")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "settings.json")

def load_settings():
    path = _settings_path()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            for k, v in DEFAULT_SETTINGS.items():
                if k not in data:
                    data[k] = v
            return data
        except Exception:
            pass
    with open(path, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_SETTINGS, f, ensure_ascii=False, indent=2)
    return dict(DEFAULT_SETTINGS)

def save_settings(settings_dict):
    with open(_settings_path(), "w", encoding="utf-8") as f:
        json.dump(settings_dict, f, ensure_ascii=False, indent=2)

settings = load_settings()

def archive_file(path):
    folder = os.path.dirname(path)
    archive = os.path.join(folder, "archive")
    os.makedirs(archive, exist_ok=True)
    shutil.copy2(path, os.path.join(archive, os.path.basename(path)))

def _bash_like_replacements(text):
    text = text.replace("Good Take", "Selected")
    text = text.replace("Reviewers Notes", "Comments DIT")
    text = text.replace("VA1A2A3A4", "V")
    text = text.replace("VA1A2A3", "V")
    text = text.replace("VA1A2", "V")
    text = text.replace("VA1", "V")
    return text

def _strip_trailing_empty_columns(cols):
    while cols and cols[-1] == "":
        cols.pop()
    return cols

def _normalize_row_to_len(row, target_len):
    if len(row) > target_len:
        return row[:target_len]
    if len(row) < target_len:
        row.extend([""] * (target_len - len(row)))
    return row

def is_selected(value):
    return str(value).strip().upper() in ("1", "Y", "YES", "TRUE", "SELECTED", "*")

def get_first_existing(row, idx_map, possible_names):
    for name in possible_names:
        if name in idx_map:
            i = idx_map[name]
            if 0 <= i < len(row):
                return row[i]
    return ""

def _safe_strip(v):
    return (v or "").strip()

def build_slate_base(scene, shot, take, fmt):
    scene = _safe_strip(scene)
    shot = _safe_strip(shot)
    take = _safe_strip(take)

    if fmt == "scene_shot_take":
        base = scene
        if shot:
            base = (base + "/" + shot) if base else shot
        if take:
            base = (base + "-" + take) if base else take
        return base.strip()

    if fmt == "scene_take_slash":
        base = scene
        if take:
            base = (base + "/" + take) if base else take
        return base.strip()

    if fmt == "scene_take_dash":
        base = scene
        if take:
            base = (base + "-" + take) if base else take
        return base.strip()

    base = scene
    if shot:
        base = (base + "/" + shot) if base else shot
    if take:
        base = (base + "-" + take) if base else take
    return base.strip()

def build_slate(scene, shot, take, selected, cam, fmt, add_selected, add_camera):
    base = build_slate_base(scene, shot, take, fmt)
    cam = _safe_strip(cam)
    star = "*" if (add_selected and is_selected(selected)) else ""

    if add_camera and cam:
        if star:
            return (base + star + " " + cam).strip()
        return (base + " " + cam).strip()

    if star:
        return (base + star).strip()

    return base.strip()

def force_tracks_column_to_v(text):
    lines = text.splitlines(True)
    out = []
    in_column = False
    in_data = False
    col_names = None
    idx = {}

    for line in lines:
        raw = line.rstrip("\n").rstrip("\r")
        st = raw.strip()

        if st == "Column":
            in_column = True
            in_data = False
            col_names = None
            idx = {}
            out.append(line)
            continue

        if st == "Data":
            in_data = True
            in_column = False
            out.append(line)
            continue

        if in_column and col_names is None and st:
            col_names = _strip_trailing_empty_columns(raw.split("\t"))
            idx = {name: i for i, name in enumerate(col_names)}
            out.append("\t".join(col_names) + "\n")
            continue

        if in_data and st and col_names and "Tracks" in idx:
            row = _strip_trailing_empty_columns(raw.split("\t"))
            row = _normalize_row_to_len(row, len(col_names))
            row[idx["Tracks"]] = "V"
            out.append("\t".join(row) + "\n")
            continue

        out.append(line)

    return "".join(out)

def unc_clipname_no_ext(source_path):
    source_path = (source_path or "").strip()
    if not source_path:
        return ""
    p = source_path.replace("\\", "/")
    base = os.path.basename(p)
    return os.path.splitext(base)[0]

def apply_postprocess_rows(text):
    lines = text.splitlines(True)
    out = []

    in_column = False
    in_data = False
    col_names = None
    idx = {}

    for line in lines:
        raw = line.rstrip("\n").rstrip("\r")
        st = raw.strip()

        if st == "Column":
            in_column = True
            in_data = False
            col_names = None
            idx = {}
            out.append(line)
            continue

        if st == "Data":
            in_data = True
            in_column = False
            out.append(line)
            continue

        if in_column and col_names is None and st:
            col_names = _strip_trailing_empty_columns(raw.split("\t"))
            idx = {name: i for i, name in enumerate(col_names)}

            if settings.get("merge_slate"):
                if "Slate" not in idx:
                    col_names.append("Slate")
                    idx["Slate"] = len(col_names) - 1

            out.append("\t".join(col_names) + "\n")
            continue

        if in_data and st and col_names:
            row = _strip_trailing_empty_columns(raw.split("\t"))
            row = _normalize_row_to_len(row, len(col_names))

            if "Tape" in idx:
                tape_mode = settings.get("tape_mode", "as_set_in_project")

                if tape_mode == "empty":
                    row[idx["Tape"]] = ""

                elif tape_mode == "reel":
                    reel_name = get_first_existing(row, idx, ["Reel", "Reel Name"])
                    row[idx["Tape"]] = _safe_strip(reel_name)

                elif tape_mode == "unc_clipname_no_ext":
                    source_path = get_first_existing(
                        row, idx,
                        ["UNC", "File Path", "Source File", "Source File Name (Full)", "Filename", "File Name", "Source Path"]
                    )
                    row[idx["Tape"]] = unc_clipname_no_ext(source_path)

                else:
                    pass

            if settings.get("merge_slate"):
                scene = get_first_existing(row, idx, ["Scene"])
                shot = get_first_existing(row, idx, ["Shot"])
                take = get_first_existing(row, idx, ["Take"])
                selected = get_first_existing(row, idx, ["Selected", "Good Take"])
                cam = get_first_existing(row, idx, ["Cam #", "Cam", "Camera", "Camera #", "Cam#", "Camera#"])

                slate_value = build_slate(
                    scene=scene,
                    shot=shot,
                    take=take,
                    selected=selected,
                    cam=cam,
                    fmt=settings.get("slate_format", "scene_shot_take"),
                    add_selected=bool(settings.get("slate_add_selected", True)),
                    add_camera=bool(settings.get("slate_add_camera", True)),
                )

                if "Slate" in idx:
                    row[idx["Slate"]] = slate_value

                if settings.get("name_mode") == "slate" and "Name" in idx:
                    row[idx["Name"]] = slate_value

            out.append("\t".join(row) + "\n")
            continue

        out.append(line)

    return "".join(out)

def format_ale_file(ale_path):
    with open(ale_path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()

    text = _bash_like_replacements(text)

    if settings.get("tracks_mode") == "video_only":
        text = force_tracks_column_to_v(text)

    text = apply_postprocess_rows(text)

    with open(ale_path, "w", encoding="utf-8") as f:
        f.write(text)

def _mpitem_unique_id(mpitem):
    try:
        return mpitem.GetUniqueId()
    except Exception:
        return None

def collect_used_mediapool_items(timeline, include_audio):
    used = {}
    if not timeline:
        return []

    for track_type in ("video", "audio"):
        if track_type == "audio" and not include_audio:
            continue

        try:
            track_count = int(timeline.GetTrackCount(track_type))
        except Exception:
            track_count = 0

        for ti in range(1, track_count + 1):
            try:
                items = timeline.GetItemListInTrack(track_type, ti) or []
            except Exception:
                items = []

            for titem in items:
                try:
                    mp = titem.GetMediaPoolItem()
                except Exception:
                    mp = None

                if not mp:
                    continue

                uid = _mpitem_unique_id(mp) or str(id(mp))
                used[uid] = mp

    return list(used.values())

def export_timeline_metadata_csv(project, timeline, out_csv_path, include_audio):
    media_pool = project.GetMediaPool()
    if not media_pool:
        return False, 0

    clips = collect_used_mediapool_items(timeline, include_audio=include_audio)
    try:
        ok = media_pool.ExportMetadata(out_csv_path, clips)
    except Exception:
        ok = False
    return bool(ok), len(clips)

def create_window():
    winID = "com.blackmagicdesign.resolve.ale_export"
    existing = ui.FindWindow(winID)
    if existing:
        existing.Show()
        existing.Raise()
        return existing

    width = 640
    height = 430

    window = dispatcher.AddWindow(
        {
            "ID": winID,
            "WindowTitle": "Export ALE and CSV",
            "WindowFlags": {"Dialog": True},
            "WindowModality": "ApplicationModal",
            "FixedSize": [width, height],
            "MinimumSize": [width, height],
            "MaximumSize": [width, height],
            "Events": {"Close": True}
        },
        ui.VGroup(
            {"Spacing": 10},
            [
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Export to", "MinimumSize": [200, 0]}),
                        ui.LineEdit({"ID": "ExportPath", "Text": settings.get("export_to", "")}),
                        ui.Button({"ID": "Browse", "Text": "Browse"})
                    ]
                ),
                ui.CheckBox({
                    "ID": "CreateTimelineFolder",
                    "Text": "Create folder with timeline name",
                    "Checked": bool(settings.get("create_export_folder_timeline_name", True))
                }),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Export format", "MinimumSize": [200, 0]}),
                        ui.ComboBox({"ID": "ExportFormatCombo"})
                    ]
                ),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Tracks", "MinimumSize": [200, 0]}),
                        ui.ComboBox({"ID": "TracksCombo"})
                    ]
                ),
                ui.CheckBox({
                    "ID": "MergeSlate",
                    "Text": "Merge slating info into Slate",
                    "Checked": bool(settings.get("merge_slate", False))
                }),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Slate format", "MinimumSize": [200, 0]}),
                        ui.ComboBox({"ID": "SlateFormatCombo"})
                    ]
                ),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Slate selected", "MinimumSize": [200, 0]}),
                        ui.CheckBox({"ID": "SlateAddSelected", "Text": "Add star if selected", "Checked": bool(settings.get("slate_add_selected", True))})
                    ]
                ),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Slate camera", "MinimumSize": [200, 0]}),
                        ui.CheckBox({"ID": "SlateAddCamera", "Text": "Add camera", "Checked": bool(settings.get("slate_add_camera", True))})
                    ]
                ),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Name column", "MinimumSize": [200, 0]}),
                        ui.ComboBox({"ID": "NameModeCombo"})
                    ]
                ),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.Label({"Text": "Tape column", "MinimumSize": [200, 0]}),
                        ui.ComboBox({"ID": "TapeModeCombo"})
                    ]
                ),
                ui.HGroup(
                    {"Spacing": 10},
                    [
                        ui.HGap(1),
                        ui.Button({"ID": "Cancel", "Text": "Cancel"}),
                        ui.Button({"ID": "Start", "Text": "Start", "Default": True})
                    ]
                )
            ]
        )
    )

    items = window.GetItems()

    items["ExportFormatCombo"].AddItem("ALE")
    items["ExportFormatCombo"].AddItem("CSV Edit Index")
    items["ExportFormatCombo"].AddItem("CSV Metadata (timeline clips)")
    items["ExportFormatCombo"].AddItem("ALE + CSV Metadata (timeline clips)")
    ef = settings.get("export_format", "ale")
    if ef == "ale":
        items["ExportFormatCombo"].CurrentIndex = 0
    elif ef == "csv_edit_index":
        items["ExportFormatCombo"].CurrentIndex = 1
    elif ef == "csv_metadata":
        items["ExportFormatCombo"].CurrentIndex = 2
    else:
        items["ExportFormatCombo"].CurrentIndex = 3

    items["TracksCombo"].AddItem("Video only")
    items["TracksCombo"].AddItem("Video + audio")
    items["TracksCombo"].CurrentIndex = 0 if settings.get("tracks_mode") == "video_only" else 1

    items["SlateFormatCombo"].AddItem("Scene/Shot-Take")
    items["SlateFormatCombo"].AddItem("Scene/Take")
    items["SlateFormatCombo"].AddItem("Scene-Take")
    fmt = settings.get("slate_format", "scene_shot_take")
    if fmt == "scene_shot_take":
        items["SlateFormatCombo"].CurrentIndex = 0
    elif fmt == "scene_take_slash":
        items["SlateFormatCombo"].CurrentIndex = 1
    else:
        items["SlateFormatCombo"].CurrentIndex = 2

    items["NameModeCombo"].AddItem("Keep clip name")
    items["NameModeCombo"].AddItem("Replace name by slate")
    items["NameModeCombo"].CurrentIndex = 0 if settings.get("name_mode") == "clipname" else 1

    items["TapeModeCombo"].AddItem("As set in project")
    items["TapeModeCombo"].AddItem("Reel")
    items["TapeModeCombo"].AddItem("UNC clipname without extension")
    items["TapeModeCombo"].AddItem("Empty")
    tape_mode = settings.get("tape_mode", "as_set_in_project")
    if tape_mode == "as_set_in_project":
        items["TapeModeCombo"].CurrentIndex = 0
    elif tape_mode == "reel":
        items["TapeModeCombo"].CurrentIndex = 1
    elif tape_mode == "unc_clipname_no_ext":
        items["TapeModeCombo"].CurrentIndex = 2
    else:
        items["TapeModeCombo"].CurrentIndex = 3

    def OnBrowse(ev):
        d = fusion.RequestDir(items["ExportPath"].Text)
        if d:
            items["ExportPath"].Text = d

    def OnCancel(ev):
        dispatcher.ExitLoop(False)

    def OnStart(ev):
        settings["export_to"] = items["ExportPath"].Text
        settings["create_export_folder_timeline_name"] = bool(items["CreateTimelineFolder"].Checked)

        ef_i = items["ExportFormatCombo"].CurrentIndex
        if ef_i == 0:
            settings["export_format"] = "ale"
        elif ef_i == 1:
            settings["export_format"] = "csv_edit_index"
        elif ef_i == 2:
            settings["export_format"] = "csv_metadata"
        else:
            settings["export_format"] = "ale_plus_csv_metadata"

        settings["tracks_mode"] = "video_only" if items["TracksCombo"].CurrentIndex == 0 else "video_audio"

        settings["merge_slate"] = bool(items["MergeSlate"].Checked)

        sf = items["SlateFormatCombo"].CurrentIndex
        if sf == 0:
            settings["slate_format"] = "scene_shot_take"
        elif sf == 1:
            settings["slate_format"] = "scene_take_slash"
        else:
            settings["slate_format"] = "scene_take_dash"

        settings["slate_add_selected"] = bool(items["SlateAddSelected"].Checked)
        settings["slate_add_camera"] = bool(items["SlateAddCamera"].Checked)

        settings["name_mode"] = "clipname" if items["NameModeCombo"].CurrentIndex == 0 else "slate"

        tm = items["TapeModeCombo"].CurrentIndex
        if tm == 0:
            settings["tape_mode"] = "as_set_in_project"
        elif tm == 1:
            settings["tape_mode"] = "reel"
        elif tm == 2:
            settings["tape_mode"] = "unc_clipname_no_ext"
        else:
            settings["tape_mode"] = "empty"

        save_settings(settings)
        dispatcher.ExitLoop(True)

    window.On["Browse"].Clicked = OnBrowse
    window.On["Cancel"].Clicked = OnCancel
    window.On["Start"].Clicked = OnStart
    window.On[winID].Close = OnCancel

    return window

project = resolve.GetProjectManager().GetCurrentProject()
assert project, "No current project"

timeline = project.GetCurrentTimeline()
assert timeline, "No current timeline"

window = create_window()
window.Show()
run = dispatcher.RunLoop()
window.Hide()

if not run:
    sys.exit(0)

export_dir = settings.get("export_to", "")
assert export_dir, "Export path is empty"

if settings.get("create_export_folder_timeline_name"):
    export_dir = os.path.join(export_dir, timeline.GetName().replace(" ", "_"))

os.makedirs(export_dir, exist_ok=True)

base_name = timeline.GetName().replace(" ", "_")
ale_path = os.path.join(export_dir, base_name + ".ale")
csv_path = os.path.join(export_dir, base_name + ".csv")
csv_meta_path = os.path.join(export_dir, base_name + "_metadata.csv")

include_audio = (settings.get("tracks_mode") != "video_only")
export_format = settings.get("export_format", "ale")

if export_format in ("ale", "ale_plus_csv_metadata"):
    if not hasattr(resolve, "EXPORT_ALE"):
        raise RuntimeError("EXPORT_ALE not available")
    ok = timeline.Export(ale_path, resolve.EXPORT_ALE, resolve.EXPORT_NONE)
    if not ok:
        raise RuntimeError("ALE export failed")
    archive_file(ale_path)
    format_ale_file(ale_path)

if export_format == "csv_edit_index":
    if not hasattr(resolve, "EXPORT_TEXT_CSV"):
        raise RuntimeError("EXPORT_TEXT_CSV not available")
    ok = timeline.Export(csv_path, resolve.EXPORT_TEXT_CSV, resolve.EXPORT_NONE)
    if not ok:
        raise RuntimeError("CSV Edit Index export failed")
    archive_file(csv_path)

if export_format in ("csv_metadata", "ale_plus_csv_metadata"):
    ok, n = export_timeline_metadata_csv(project, timeline, csv_meta_path, include_audio=include_audio)
    if not ok:
        raise RuntimeError("CSV Metadata export failed")
    archive_file(csv_meta_path)
    print("CSV Metadata clips exported:", n)

print("Export done")