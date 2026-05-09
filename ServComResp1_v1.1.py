#!/usr/bin/env python3
# coding: utf-8
"""
Détecte changements d'angle d'une courbe verte et classe.
Lit steam_games_cards_prices.xlsx, colonne image (fallback),
traite images dans canvas_images et écrit XLSX.
Version : n'écrit pas les colonnes nb_angles et error.
"""

import os
import math
import glob
import shutil
import traceback
import numpy as np
import pandas as pd
import cv2

# ---------- Configuration ----------
XLSX_INPUT = "steam_games_cards_prices.xlsx"
SHEET = 0
IMG_COL_NAME = "chemin image canvas"
IMAGES_DIR = r"C:\Users\Megaport\canvas_images"
XLSX_OUT = "steam_games_cards_frequences.xlsx"

GREEN_LO = np.array([40, 60, 40])
GREEN_HI = np.array([90, 255, 255])

CROP_LEFT = 0.02
CROP_RIGHT = 0.02
CROP_TOP = 0.05
CROP_BOTTOM = 0.10

SMOOTH_KERNEL = 11
ANGLE_CHANGE_DEG = 5.0
MIN_SEP_PIX = 5

DEBUG_DIR = "debug_bad_images"
os.makedirs(DEBUG_DIR, exist_ok=True)

GAME_COL_CANDS = [
    "game", "game_name", "nom du jeu", "nom_du_jeu",
    "jeu", "title", "app", "app_name"
]
CARD_COL_CANDS = [
    "card", "card_name", "nom de la carte",
    "nom_de_la_carte", "carte", "item"
]


def find_col(cols, cands):
    cols_l = {c.lower(): c for c in cols}
    for cand in cands:
        if cand in cols:
            return cand
        if cand.lower() in cols_l:
            return cols_l[cand.lower()]
    return None


def load_img(path):
    """Lecture robuste: BGRA->BGR, gray->BGR."""
    img = cv2.imread(path, cv2.IMREAD_UNCHANGED)
    if img is None:
        raise FileNotFoundError("Image introuvable: " + path)
    if not hasattr(img, "shape"):
        raise ValueError("Image sans shape: " + path)
    if img.ndim == 3 and img.shape[2] == 4:
        try:
            img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
        except Exception:
            img = img[:, :, :3].copy()
    if img.ndim == 2:
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
    if img.ndim != 3 or img.shape[2] != 3:
        raise ValueError("Image attendue (H,W,3), reçu: " +
                         str(img.shape))
    return img


def crop_image(img):
    h, w = img.shape[:2]
    left_px = int(w * CROP_LEFT)
    right_px = int(w * (1 - CROP_RIGHT))
    top_px = int(h * CROP_TOP)
    bottom_px = int(h * (1 - CROP_BOTTOM))
    if right_px <= left_px or bottom_px <= top_px:
        return img
    return img[top_px:bottom_px, left_px:right_px]


def resolve_image_path(img_dir, img_rel_s):
    if img_rel_s is None:
        return ""
    p = str(img_rel_s).strip().strip('"\'')

    if not p:
        return ""

    p = p.replace("\\", "/")

    if os.path.isabs(p):
        if os.path.isfile(p):
            return p
        d, fname = os.path.split(p)
        if d and os.path.isdir(d):
            pattern = os.path.join(d, fname)
            matches = glob.glob(pattern)
            if matches:
                return matches[0]
        return p

    cand = os.path.normpath(os.path.join(img_dir, p))
    if os.path.isfile(cand):
        return cand

    base, ext = os.path.splitext(cand)
    if ext:
        dirn = os.path.dirname(base)
        name = os.path.basename(base)
        pattern = os.path.join(dirn, name + ".*")
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    else:
        for e in (".png", ".jpg", ".jpeg", ".bmp", ".tif",
                  ".tiff"):
            cand2 = base + e
            if os.path.isfile(cand2):
                return cand2

    fname = os.path.basename(p)
    pattern = os.path.join(img_dir, "**", fname)
    matches = glob.glob(pattern, recursive=True)
    if matches:
        return matches[0]

    return cand


def mask_green(img_bgr):
    if img_bgr is None:
        raise ValueError("mask_green reçu None")
    if img_bgr.ndim != 3 or img_bgr.shape[2] != 3:
        msg = ("mask_green attend (H,W,3), reçu: " +
               str(getattr(img_bgr, "shape", None)))
        raise ValueError(msg)
    hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, GREEN_LO, GREEN_HI)
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (3, 3))
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k,
                            iterations=1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, k,
                            iterations=1)
    return mask


def extract_curve_y_by_x(mask):
    h, w = mask.shape[:2]
    ys = np.full(w, np.nan, dtype=float)
    for x in range(w):
        rows = np.where(mask[:, x] > 0)[0]
        if rows.size:
            ys[x] = rows.mean()
    return ys, h


def _gaussian_kernel_1d(k):
    if k <= 1:
        return np.array([1.0], dtype=float)
    if k % 2 == 0:
        k += 1
    center = (k - 1) / 2.0
    sigma = max(0.5, k / 3.0)
    x = np.arange(k, dtype=float)
    g = np.exp(-0.5 * ((x - center) / sigma) ** 2)
    g /= g.sum()
    return g


def fill_and_smooth(y, h):
    w = len(y)
    if np.isnan(y).all():
        y = np.full(w, h / 2.0, dtype=float)
    elif np.isnan(y).any():
        good = ~np.isnan(y)
        if good.sum() >= 2:
            x_good = np.flatnonzero(good)
            y_good = y[good]
            x_all = np.arange(w)
            y = np.interp(x_all, x_good, y_good)
        else:
            y = np.full(w, h / 2.0, dtype=float)

    k = SMOOTH_KERNEL if SMOOTH_KERNEL % 2 == 1 else (
        SMOOTH_KERNEL + 1)
    k = min(k, max(3, w if w % 2 == 1 else w - 1))
    if k < 3:
        k = 3 if w >= 3 else 1

    kernel = _gaussian_kernel_1d(k)
    y_blur = np.convolve(y, kernel, mode="same")
    return y_blur


def compute_angle_changes(y):
    dy = np.diff(y)
    dx = 1.0
    angles = np.arctan2(dy, dx)
    dtheta = np.diff(angles)
    dtheta = (dtheta + np.pi) % (2 * np.pi) - np.pi
    ddeg = np.abs(np.degrees(dtheta))
    idxs = np.where(ddeg >= ANGLE_CHANGE_DEG)[0]
    if idxs.size == 0:
        return 0, ddeg
    events = 0
    last = -9999
    for i in idxs:
        if i - last >= MIN_SEP_PIX:
            events += 1
            last = i
    return events, ddeg


def map_count_to_class(count):
    if count == 0:
        return 0
    if count < 2:
        return 1
    if 2 <= count <= 10:
        return 2
    cls = 2 + int(math.ceil((count - 10) / 10.0))
    return min(cls, 10)


def analyze_and_write(xlsx_in=XLSX_INPUT, sheet=SHEET,
                      img_col_name=IMG_COL_NAME,
                      img_dir=IMAGES_DIR, xlsx_out=XLSX_OUT):
    df = pd.read_excel(xlsx_in, sheet_name=sheet,
                       engine="openpyxl")
    cols = list(df.columns)
    if img_col_name in cols:
        img_col = img_col_name
    else:
        found = find_col(cols, [img_col_name])
        if found:
            img_col = found
        else:
            if len(cols) >= 7:
                img_col = cols[6]
            else:
                raise KeyError("Colonne image introuvable.")
    game_col = find_col(cols, GAME_COL_CANDS)
    card_col = find_col(cols, CARD_COL_CANDS)
    if game_col is None and len(cols) >= 1:
        game_col = cols[0]
    if card_col is None and len(cols) >= 2:
        card_col = cols[1]

    out_rows = []

    for idx, row in df.iterrows():
        game = row.get(game_col, "") if game_col in df.columns else ""
        card = row.get(card_col, "") if card_col in df.columns else ""
        img_rel = row.get(img_col, "")
        img_rel_s = "" if pd.isna(img_rel) else str(img_rel)
        img_path = resolve_image_path(img_dir, img_rel_s)
        entry = {
            "Nom du jeu": game,
            "Nom de la carte": card,
            "Chemin image (input)": img_rel_s,
            "Classe fréquence (angles)": None,
            "nb_angles": None,
            "error": ""
        }
        if not img_rel_s:
            entry["error"] = (
                "no_image_path; cell=" + repr(img_rel) +
                "; cwd=" + os.getcwd()
            )
            entry["Classe fréquence (angles)"] = 0
            entry["nb_angles"] = 0
            out_rows.append(entry)
            continue
        if not img_path or not os.path.isfile(img_path):
            entry["error"] = (
                "file_not_found; tried=" + repr(img_path) +
                "; cell=" + repr(img_rel) +
                "; cwd=" + os.getcwd()
            )
            entry["Classe fréquence (angles)"] = 0
            entry["nb_angles"] = 0
            out_rows.append(entry)
            continue
        try:
            img = load_img(img_path)
            entry["error"] = (
                entry.get("error", "") +
                f" debug_shape={img.shape}, dtype={img.dtype}"
            )
            img_c = crop_image(img)
            if img_c.size == 0 or img_c.shape[0] < 2 or \
               img_c.shape[1] < 2:
                raise ValueError("Image crop trop petite: " +
                                 str(img_c.shape))
            mask = mask_green(img_c)
            y_raw, h = extract_curve_y_by_x(mask)
            y = fill_and_smooth(y_raw, h)
            nb_angles, ddeg = compute_angle_changes(y)
            cls = map_count_to_class(nb_angles)
            entry["nb_angles"] = int(nb_angles)
            entry["Classe fréquence (angles)"] = int(cls)
            out_rows.append(entry)
        except Exception as e:
            tb = traceback.format_exc()
            try:
                bad_name = ("bad_idx" + str(idx) + "_" +
                            os.path.basename(img_path))
                bad_path = os.path.join(DEBUG_DIR, bad_name)
                if "img" in locals() and isinstance(img, np.ndarray):
                    cv2.imwrite(bad_path, img)
                else:
                    if os.path.isfile(img_path):
                        shutil.copy2(img_path, bad_path)
            except Exception:
                pass
            entry["error"] = (
                "proc_error: " + str(e) + "; path=" + img_path +
                "; traceback=" + tb
            )
            entry["Classe fréquence (angles)"] = 0
            entry["nb_angles"] = 0
            out_rows.append(entry)

    full_df = pd.DataFrame(
        out_rows,
        columns=[
            "Nom du jeu", "Nom de la carte",
            "Chemin image (input)",
            "Classe fréquence (angles)", "nb_angles", "error"
        ],
    )

    export_cols = [
        "Nom du jeu", "Nom de la carte",
        "Chemin image (input)",
        "Classe fréquence (angles)"
    ]
    out_df = full_df[export_cols]

    out_df.to_excel(xlsx_out, index=False, engine="openpyxl")
    print("Terminé. Résultats écrits dans:", xlsx_out)


if __name__ == "__main__":
    analyze_and_write()
