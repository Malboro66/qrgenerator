from __future__ import annotations

import tkinter as tk
from typing import Callable, Optional

try:
    from PIL import Image, ImageTk

    _PIL_OK = True
except ImportError:
    _PIL_OK = False


class PreviewInterativo(tk.Canvas):
    """Canvas interativo para preview de etiquetas com drag-to-resize."""

    HANDLE_R: int = 6
    MIN_CODE_MM: float = 5.0
    MIN_LABEL_MM: float = 20.0

    _C_CODE = "#2563eb"
    _C_LBL = "#6b7280"
    _C_BG = "#d5d3cc"

    def __init__(
        self,
        parent,
        on_code_resized: Callable[[float, float], None],
        on_label_resized: Callable[[float, float], None],
        **kwargs,
    ) -> None:
        kwargs.setdefault("bg", self._C_BG)
        kwargs.setdefault("highlightthickness", 0)
        super().__init__(parent, **kwargs)

        self._on_code = on_code_resized
        self._on_lbl = on_label_resized

        self._lw: float = 100.0
        self._lh: float = 60.0
        self._cx: float = 6.0
        self._cy: float = 6.0
        self._cw: float = 38.0
        self._ch: float = 38.0

        self._scale: float = 1.0
        self._origin: tuple[float, float] = (0.0, 0.0)

        self._drag: Optional[dict] = None
        self._handles: list[dict] = []
        self._tk_img = None
        self._code_img = None

        self.bind("<ButtonPress-1>", self._press)
        self.bind("<B1-Motion>", self._motion)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Motion>", self._hover)
        self.bind("<Configure>", lambda _e: self.after_idle(self._draw))

    def atualizar(
        self,
        lbl_w_mm: float,
        lbl_h_mm: float,
        cod_w_mm: float,
        cod_h_mm: float,
        codigo_img=None,
    ) -> None:
        self._lw = max(self.MIN_LABEL_MM, float(lbl_w_mm))
        self._lh = max(self.MIN_LABEL_MM, float(lbl_h_mm))
        self._cw = max(self.MIN_CODE_MM, min(float(cod_w_mm), self._lw - self._cx - 2))
        self._ch = max(self.MIN_CODE_MM, min(float(cod_h_mm), self._lh - self._cy - 2))
        self._cx = min(self._cx, self._lw - self._cw - 2)
        self._cy = min(self._cy, self._lh - self._ch - 2)
        self._code_img = codigo_img
        self._draw()

    def mover_codigo(self, x_mm: float, y_mm: float) -> None:
        self._cx = max(1.0, min(float(x_mm), self._lw - self._cw - 1))
        self._cy = max(1.0, min(float(y_mm), self._lh - self._ch - 1))
        self._draw()

    def obter_estado(self) -> dict:
        return {
            "etiqueta_w_mm": round(self._lw, 1),
            "etiqueta_h_mm": round(self._lh, 1),
            "codigo_w_mm": round(self._cw, 1),
            "codigo_h_mm": round(self._ch, 1),
            "codigo_x_mm": round(self._cx, 1),
            "codigo_y_mm": round(self._cy, 1),
        }

    def _get_scale(self) -> float:
        cw = self.winfo_width() or 500
        ch = self.winfo_height() or 360
        pad = 54
        if self._lw <= 0 or self._lh <= 0:
            return 1.0
        return min((cw - pad * 2) / self._lw, (ch - pad * 2) / self._lh)

    def _lbl2px(self, xmm: float, ymm: float) -> tuple[float, float]:
        ox, oy = self._origin
        s = self._scale
        return ox + xmm * s, oy + ymm * s

    def _draw(self) -> None:
        cw = self.winfo_width() or 500
        ch = self.winfo_height() or 360
        self.delete("all")
        self._handles = []

        s = self._get_scale()
        self._scale = s
        lw_px = self._lw * s
        lh_px = self._lh * s
        ox = (cw - lw_px) / 2
        oy = (ch - lh_px) / 2
        self._origin = (ox, oy)

        self.create_rectangle(0, 0, cw, ch, fill=self._C_BG, outline="")

        self.create_rectangle(
            ox + 3,
            oy + 4,
            ox + lw_px + 3,
            oy + lh_px + 4,
            fill="#b8b5ad",
            outline="",
        )
        self.create_rectangle(
            ox,
            oy,
            ox + lw_px,
            oy + lh_px,
            fill="white",
            outline="#c4c9d0",
            width=1,
        )

        cpx, cpy = self._lbl2px(self._cx, self._cy)
        cw_px = self._cw * s
        ch_px = self._ch * s

        if _PIL_OK and self._code_img is not None:
            try:
                img_r = self._code_img.resize(
                    (max(1, int(cw_px)), max(1, int(ch_px))),
                    Image.Resampling.LANCZOS,
                )
                self._tk_img = ImageTk.PhotoImage(img_r)
                self.create_image(cpx, cpy, image=self._tk_img, anchor="nw")
            except Exception:
                self._draw_placeholder(cpx, cpy, cw_px, ch_px)
        else:
            self._draw_placeholder(cpx, cpy, cw_px, ch_px)

        self.create_rectangle(
            cpx - 1,
            cpy - 1,
            cpx + cw_px + 1,
            cpy + ch_px + 1,
            fill="",
            outline=self._C_CODE,
            width=1.5,
            dash=(4, 2),
        )

        for hid, hx, hy in [
            ("c-nw", cpx, cpy),
            ("c-ne", cpx + cw_px, cpy),
            ("c-se", cpx + cw_px, cpy + ch_px),
            ("c-sw", cpx, cpy + ch_px),
        ]:
            self._handle(hid, hx, hy, self._C_CODE)

        for hid, hx, hy in [
            ("l-e", ox + lw_px, oy + lh_px / 2),
            ("l-s", ox + lw_px / 2, oy + lh_px),
            ("l-se", ox + lw_px, oy + lh_px),
        ]:
            self._handle(hid, hx, hy, self._C_LBL)

        self._dim(ox, oy - 22, ox + lw_px, oy - 22, f"{self._lw:.0f} mm", self._C_LBL)
        self._dim(ox - 28, oy, ox - 28, oy + lh_px, f"{self._lh:.0f} mm", self._C_LBL)
        if cw_px > 32:
            self._dim(cpx, cpy + ch_px + 17, cpx + cw_px, cpy + ch_px + 17, f"{self._cw:.0f} mm", self._C_CODE)
        if ch_px > 32:
            self._dim(cpx + cw_px + 17, cpy, cpx + cw_px + 17, cpy + ch_px, f"{self._ch:.0f} mm", self._C_CODE)

    def _draw_placeholder(self, x: float, y: float, w: float, h: float) -> None:
        self.create_rectangle(x, y, x + w, y + h, fill="#f8fafc", outline="")
        n = max(3, min(15, int(w / 7)))
        if n < 1:
            return
        cw_cell = w / n
        ch_cell = h / n
        for r in range(n):
            for c in range(n):
                in_tl = r < 3 and c < 3
                in_tr = r < 3 and c >= n - 3
                in_bl = r >= n - 3 and c < 3
                is_d = ((r ^ c) * (r + c + 1)) % 7 < 3
                if in_tl or in_tr or in_bl or is_d:
                    self.create_rectangle(
                        x + c * cw_cell + 0.5,
                        y + r * ch_cell + 0.5,
                        x + (c + 1) * cw_cell - 0.5,
                        y + (r + 1) * ch_cell - 0.5,
                        fill="#111827",
                        outline="",
                    )

    def _handle(self, hid: str, hx: float, hy: float, color: str) -> None:
        r = self.HANDLE_R
        active = self._drag and self._drag["id"] == hid
        self.create_oval(
            hx - r,
            hy - r,
            hx + r,
            hy + r,
            fill=color if active else "white",
            outline=color,
            width=2,
        )
        self._handles.append({"id": hid, "x": hx, "y": hy})

    def _dim(self, x1: float, y1: float, x2: float, y2: float, text: str, color: str) -> None:
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        self.create_line(x1, y1, x2, y2, fill=color, width=1, dash=(3, 2))
        tw = max(30, len(text) * 6)
        self.create_rectangle(
            mx - tw / 2 - 2,
            my - 7,
            mx + tw / 2 + 2,
            my + 7,
            fill=self._C_BG,
            outline="",
        )
        self.create_text(mx, my, text=text, fill=color, font=("Segoe UI", 8), anchor="center")

    def _hit(self, x: float, y: float) -> Optional[dict]:
        tol = self.HANDLE_R + 6
        for h in self._handles:
            if (x - h["x"]) ** 2 + (y - h["y"]) ** 2 <= tol * tol:
                return h
        return None

    _CURSORS = {
        "c-nw": "crosshair",
        "c-ne": "crosshair",
        "c-se": "crosshair",
        "c-sw": "crosshair",
        "l-e": "crosshair",
        "l-s": "crosshair",
        "l-se": "crosshair",
    }

    def _press(self, e: tk.Event) -> None:
        h = self._hit(e.x, e.y)
        if not h:
            return
        self._drag = {
            "id": h["id"],
            "sx": e.x,
            "sy": e.y,
            "sc": self._scale,
            "olw": self._lw,
            "olh": self._lh,
            "ocx": self._cx,
            "ocy": self._cy,
            "ocw": self._cw,
            "och": self._ch,
        }

    def _motion(self, e: tk.Event) -> None:
        if not self._drag:
            return
        d = self._drag
        s = d["sc"]
        dx = (e.x - d["sx"]) / s
        dy = (e.y - d["sy"]) / s
        hid = d["id"]
        mc, ml = self.MIN_CODE_MM, self.MIN_LABEL_MM

        if hid.startswith("c"):
            x, y, w, h = d["ocx"], d["ocy"], d["ocw"], d["och"]
            if hid == "c-se":
                w = max(mc, w + dx)
                h = max(mc, h + dy)
            elif hid == "c-sw":
                x += dx
                w = max(mc, w - dx)
                h = max(mc, h + dy)
            elif hid == "c-ne":
                w = max(mc, w + dx)
                y += dy
                h = max(mc, h - dy)
            elif hid == "c-nw":
                x += dx
                y += dy
                w = max(mc, w - dx)
                h = max(mc, h - dy)
            x = max(1.0, min(x, self._lw - w - 1.0))
            y = max(1.0, min(y, self._lh - h - 1.0))
            w = min(w, self._lw - x - 1.0)
            h = min(h, self._lh - y - 1.0)
            self._cx = x
            self._cy = y
            self._cw = w
            self._ch = h
            self._on_code(round(w, 1), round(h, 1))

        elif hid.startswith("l"):
            lw, lh = d["olw"], d["olh"]
            if hid in ("l-e", "l-se"):
                lw = max(ml, lw + dx)
            if hid in ("l-s", "l-se"):
                lh = max(ml, lh + dy)
            self._lw = lw
            self._lh = lh
            self._cx = min(self._cx, lw - self._cw - 1.0)
            self._cy = min(self._cy, lh - self._ch - 1.0)
            self._on_lbl(round(lw, 1), round(lh, 1))

        self._draw()

    def _release(self, _e: tk.Event) -> None:
        self._drag = None

    def _hover(self, e: tk.Event) -> None:
        if self._drag:
            return
        h = self._hit(e.x, e.y)
        self.configure(cursor=self._CURSORS.get(h["id"], "") if h else "")
