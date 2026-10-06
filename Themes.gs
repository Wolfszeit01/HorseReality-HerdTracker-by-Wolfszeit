/* ══════════════════════════════════════════════════════════════
   THEME-DATEN
   Reine Farbdefinitionen. Keine Logik, keine Sheet-Konfiguration.
   Funktionen (getActiveTheme, switchTheme, renderThemeVariables) 
   liegen in ThemeBridge.gs
   ══════════════════════════════════════════════════════════════ */

const THEMES = {

  OCEAN_BLUE: {
    primary: {
      deep:  "#0B2B4A",
      dark:  "#1A4A6E",
      mid:   "#2E6A96",
      light: "#7BAEC6",
      ultra: "#C5E2F0",
      lightmuted: "#F0F6FA",
      ultramuted: "#E5EFF5"
    },
    secondary: {
      deep:  "#3A2E1D",
      dark:  "#6B5434",
      mid:   "#9E7A4D",
      base:  "#C4A67A",
      light: "#E2D2B6",
    },
    tertiary: {
      deep:  "#1E3A2E",
      dark:  "#2E5A45",
      mid:   "#4A7A61",
      light: "#8ABA9A",
      ultra: "#BDE0C8",
    },
    accents: {
    accent1: "#6B4A5A",
    accent2: "#8B5A42",
    accent3: "#4A6B80",
    accent4: "#3A7A80",
    accent5: "#C49A3C",
    accent6: "#B87A42",
    accent7: "#5A5248",
    },
    neutrals: {
      neutralcolor: "#F2EBEA",
      white: "#FFFFFF",
      grey: "#7A7578"
    }
  },
  FOREST_GOLD: {
  "primary": {
    "deep": "#19372C",
    "dark": "#254D3C",
    "mid": "#42694E",
    "light": "#91AD8D",
    "ultra": "#C9D9BD",
    "lightmuted": "#FAFAF3",
    "ultramuted": "#F1F3E8"
  },
  "secondary": {
    "deep": "#65501F",
    "dark": "#907139",
    "mid": "#BE9851",
    "base": "#D8BB79",
    "light": "#E9D9AD"
  },
  "tertiary": {
    "deep": "#294D38",
    "dark": "#406447",
    "mid": "#648360",
    "light": "#95AE80",
    "ultra": "#C0D0A9"
  },
  "accents": {
    "accent1": "#A77985",
    "accent2": "#A34E43",
    "accent3": "#617D91",
    "accent4": "#4E8680",
    "accent5": "#BE9848",
    "accent6": "#AB7447",
    "accent7": "#77614D"
  },
  "neutrals": {
    "neutralcolor": "#E3E8D9",
    "white": "#FFFFFF",
    "grey": "#70796E"
  }
},
  MIDNIGHT_SILVER: {
  "primary": {
    "deep": "#172338",
    "dark": "#263650",
    "mid": "#465B78",
    "light": "#94A8C2",
    "ultra": "#CDD8E7",
    "lightmuted": "#F8FAFD",
    "ultramuted": "#EEF2F8"
  },
  "secondary": {
    "deep": "#424F63",
    "dark": "#65748A",
    "mid": "#9AAAC0",
    "base": "#CCD6E5",
    "light": "#E3E9F2"
  },
  "tertiary": {
    "deep": "#2D504A",
    "dark": "#466C61",
    "mid": "#6B9080",
    "light": "#9DBBAC",
    "ultra": "#C5DACE"
  },
  "accents": {
    "accent1": "#A47F9C",
    "accent2": "#A75361",
    "accent3": "#718AAC",
    "accent4": "#558A94",
    "accent5": "#BAA06C",
    "accent6": "#B47E60",
    "accent7": "#807A8B"
  },
  "neutrals": {
    "neutralcolor": "#DFE5EF",
    "white": "#FFFFFF",
    "grey": "#727D8D"
  }
}
};

// Aktives Theme (State) — Fallback, falls PropertiesService noch nichts gespeichert hat.
// Die eigentliche Persistenz läuft über PropertiesService (siehe ThemeBridge.gs),
// da normale Variablen zwischen Skript-Ausführungen NICHT erhalten bleiben.
const DEFAULT_THEME_KEY = "OCEAN_BLUE";