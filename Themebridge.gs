/* ══════════════════════════════════════════════════════════════
   THEME BRIDGE
   Verbindet Themes.gs (Rohfarben) mit Sheets UND HTML-Modals.
   
   Kernidee: renderThemeVariables() läuft automatisch durch die 
   Theme-Struktur (primary/secondary/tertiary/accents/neutrals) 
   und erzeugt daraus CSS-Variablen wie --primary-dark, 
   --secondary-base, --accent-1 usw. Keine manuelle Liste nötig – 
   funktioniert auch, wenn die Palette später erweitert wird.
   ══════════════════════════════════════════════════════════════ */

/**
 * Liefert das aktuell aktive Theme-Objekt (für Sheets UND Modals).
 * Liest den Theme-Key aus PropertiesService, da normale Variablen
 * zwischen Skript-Ausführungen NICHT erhalten bleiben.
 */
function getActiveTheme() {
  return availableThemes_()[getActiveThemeKey()];
}

/**
 * Liefert den aktuell gespeicherten Theme-Key. Fällt auf DEFAULT_THEME_KEY
 * zurück, falls noch nie ein Theme gesetzt wurde.
 */
function getActiveThemeKey() {
  const stored = PropertiesService.getDocumentProperties().getProperty('PUBLIC_ACTIVE_THEME_KEY');
  return (stored && Object.prototype.hasOwnProperty.call(availableThemes_(), stored)) ? stored : DEFAULT_THEME_KEY;
}

/**
 * Wechselt das aktive Theme (dauerhaft gespeichert) und wendet es
 * sofort auf alle Sheets an. HTML-Modals ziehen sich das neue Theme
 * automatisch beim nächsten Öffnen (via renderThemeVariables()),
 * weil sie bei jedem Öffnen getActiveTheme() neu abfragen.
 */
function switchTheme(themeKey) {
  if (!Object.prototype.hasOwnProperty.call(availableThemes_(), themeKey)) {
    throw new Error('Theme "' + themeKey + '" nicht gefunden.');
  }
  PropertiesService.getDocumentProperties().setProperty('PUBLIC_ACTIVE_THEME_KEY', themeKey);
  applyThemeToAllSheets();
}

/**
 * Baut den <style>:root{...}</style>-Block für HTML-Modals.
 * Läuft automatisch durch das aktive Theme-Objekt und erzeugt für 
 * jeden Farbton eine CSS-Variable nach dem Schema --gruppe-ton, z.B.:
 *   theme.primary.dark    -> --primary-dark
 *   theme.accents.accent2 -> --accent-2 (Sonderfall, siehe unten)
 *   theme.neutrals.white  -> --neutrals-white
 *
 * Einbindung in jedem Modal-<head>: <?!= renderThemeVariables(); ?>
 */
function renderThemeVariables() {
  const theme = getActiveTheme();
  const lines = [];

  Object.keys(theme).forEach(function (group) {         // primary, secondary, tertiary, accents, neutrals
    Object.keys(theme[group]).forEach(function (tone) {  // deep, dark, mid, light, accent1, white...
      // Bei accents: "accent1" -> "accent-1" (lesbarer als --accents-accent1)
      const cssName = (group === 'accents')
        ? tone.replace('accent', 'accent-')
        : group + '-' + tone;
      lines.push('  --' + cssName + ': ' + theme[group][tone] + ';');
    });
  });

  return '<style>\n:root {\n' + lines.join('\n') + '\n}\n</style>';
}

/**
 * Standard-Include-Helper für HTML-Templates (Apps-Script-Pattern).
 * Nutzung im HTML: <?!= include('_Shared_ThemeStyle'); ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}