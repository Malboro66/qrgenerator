import json
import locale
import os


class I18nService:
    def __init__(self, locale_dir: str | None = None, default_locale: str = "pt_BR"):
        base_dir = os.path.dirname(os.path.dirname(__file__))
        self.locale_dir = locale_dir or os.path.join(base_dir, "locales")
        self.default_locale = default_locale
        self.current_locale = self._detect_system_locale() or default_locale
        self._cache: dict[str, dict[str, str]] = {}
        self._translations = self._load_locale(self.current_locale)

    def _detect_system_locale(self) -> str | None:
        candidates: list[str] = []
        for category in (locale.LC_ALL, locale.LC_TIME, locale.LC_MONETARY):
            try:
                loc = locale.getlocale(category)[0]
                if loc:
                    candidates.append(loc)
            except Exception:
                continue

        for loc in candidates:
            normalized = loc.replace("-", "_")
            if self._locale_file_exists(normalized):
                return normalized
            base = normalized.split("_")[0]
            if base.startswith("pt") and self._locale_file_exists("pt_BR"):
                return "pt_BR"
            if base.startswith("en") and self._locale_file_exists("en_US"):
                return "en_US"
        return None

    def _locale_file_exists(self, locale_code: str) -> bool:
        return os.path.exists(os.path.join(self.locale_dir, f"{locale_code}.json"))

    def _load_locale(self, locale_code: str) -> dict[str, str]:
        if locale_code in self._cache:
            return self._cache[locale_code]

        caminho = os.path.join(self.locale_dir, f"{locale_code}.json")
        if not os.path.exists(caminho):
            if locale_code != self.default_locale:
                return self._load_locale(self.default_locale)
            return {}

        try:
            with open(caminho, "r", encoding="utf-8") as f:
                dados = json.load(f)
                if not isinstance(dados, dict):
                    dados = {}
        except Exception:
            dados = {}

        self._cache[locale_code] = {str(k): str(v) for k, v in dados.items()}
        return self._cache[locale_code]

    def set_locale(self, locale_code: str):
        locale_code = locale_code.replace("-", "_")
        self.current_locale = locale_code
        self._translations = self._load_locale(locale_code)

    def t(self, key: str, default: str = "", **kwargs) -> str:
        value = self._translations.get(key)
        if value is None and self.current_locale != self.default_locale:
            value = self._load_locale(self.default_locale).get(key)
        texto = value if value is not None else (default or key)
        if kwargs:
            try:
                return texto.format(**kwargs)
            except Exception:
                return texto
        return texto
