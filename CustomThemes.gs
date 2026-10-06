var customThemeCache_;
function customThemes_() {
  if(customThemeCache_) return customThemeCache_;
  const all=PropertiesService.getDocumentProperties().getProperties();
  const result={};
  Object.keys(all).filter(key=>/^PUBLIC_CUSTOM_THEME_[a-f0-9-]+$/.test(key)).forEach(key=>{
    try { const record=JSON.parse(all[key]); record.colors=validateThemeColors_(record.colors); record.palette=buildCustomPalette_(record.colors); result[key]=record; } catch(error) { Logger.log('Invalid custom theme: '+key); }
  });
  return customThemeCache_=result;
}
function availableThemes_() {
  const themes=Object.assign({},THEMES);
  Object.entries(customThemes_()).forEach(([key,record])=>themes[key]=record.palette);
  return themes;
}
function validateThemeColors_(colors) {
  const clean={};
  ['main','accent','background','text','positive'].forEach(key=>{
    if(!colors || typeof colors[key]!=='string' || !/^#[a-f0-9]{6}$/i.test(colors[key])) throw new Error('Choose a valid colour for '+key+'.');
    clean[key]=colors[key].toUpperCase();
  });
  return clean;
}
function themeBlend_(a,b,weight) {
  return '#'+[1,3,5].map(i=>Math.round(parseInt(a.slice(i,i+2),16)*weight+parseInt(b.slice(i,i+2),16)*(1-weight)).toString(16).padStart(2,'0')).join('');
}
function buildCustomPalette_(input) {
  const c=validateThemeColors_(input), mix=themeBlend_;
  return {
    primary:{deep:c.text,dark:c.main,mid:mix(c.main,c.text,.7),light:mix(c.main,c.background,.4),ultra:mix(c.main,c.background,.2),lightmuted:mix(c.main,c.background,.025),ultramuted:mix(c.main,c.background,.07)},
    secondary:{deep:mix(c.accent,c.text,.45),dark:mix(c.accent,c.text,.7),mid:mix(c.accent,c.text,.9),base:c.accent,light:mix(c.accent,c.background,.45)},
    tertiary:{deep:mix(c.positive,c.text,.3),dark:mix(c.positive,c.text,.55),mid:mix(c.positive,c.text,.8),light:c.positive,ultra:mix(c.positive,c.background,.7)},
    accents:{accent1:mix(c.main,'#BD7E87',.3),accent2:'#A15849',accent3:mix(c.main,'#60798D',.3),accent4:mix(c.positive,'#4D8489',.3),accent5:'#C49A3C',accent6:'#B87A42',accent7:mix(c.text,'#6D5A4D',.3)},
    neutrals:{neutralcolor:mix(c.main,c.background,.16),white:c.background,grey:mix(c.text,c.background,.55)}
  };
}
function themeContrast_(a,b) {
  const lum=hex=>{const rgb=[1,3,5].map(i=>parseInt(hex.slice(i,i+2),16)/255).map(v=>v<=.04045?v/12.92:Math.pow((v+.055)/1.055,2.4));return rgb[0]*.2126+rgb[1]*.7152+rgb[2]*.0722;};
  const x=lum(a),y=lum(b);return (Math.max(x,y)+.05)/(Math.min(x,y)+.05);
}
function previewCustomTheme(colors) {
  const palette=buildCustomPalette_(colors);
  const pairs=[['Body text',palette.primary.deep,palette.primary.ultramuted],['Buttons',palette.neutrals.white,palette.primary.dark],['Header title',palette.secondary.base,palette.primary.dark],['Table header',palette.neutrals.white,palette.primary.deep]];
  return {palette,contrast:pairs.map(([label,fg,bg])=>({label,ratio:Math.round(themeContrast_(fg,bg)*100)/100}))};
}
function themeEditorState() {
  const labels={OCEAN_BLUE:'Ocean Blue',FOREST_GOLD:'Forest & Gold',MIDNIGHT_SILVER:'Midnight & Silver'};
  const customs=customThemes_();
  return {active:getActiveThemeKey(),themes:Object.entries(availableThemes_()).map(([id,palette])=>({id,name:customs[id]?customs[id].name:(labels[id]||id),custom:!!customs[id],colors:customs[id]?customs[id].colors:{main:palette.primary.dark,accent:palette.secondary.base,background:palette.neutrals.white,text:palette.primary.deep,positive:palette.tertiary.light},palette}))};
}
function saveAndApplyCustomTheme(input) {
  if(!input || typeof input.name!=='string' || !input.name.trim() || input.name.trim().length>40) throw new Error('Enter a theme name (1–40 characters).');
  const colors=validateThemeColors_(input.colors);
  let id=input.id;
  if(id && (!/^PUBLIC_CUSTOM_THEME_[a-f0-9-]+$/.test(id) || !customThemes_()[id])) throw new Error('This custom theme no longer exists. Reload the editor.');
  if(!id) {
    if(Object.keys(customThemes_()).length>=20) throw new Error('You can store up to 20 custom themes. Edit or delete an existing theme.');
    id='PUBLIC_CUSTOM_THEME_'+Utilities.getUuid();
  }
  PropertiesService.getDocumentProperties().setProperty(id,JSON.stringify({name:input.name.trim(),colors}));
  customThemeCache_=null;
  switchTheme(id);
  return themeEditorState();
}
function activateEditorTheme(id) { switchTheme(id); return themeEditorState(); }
function deleteCustomTheme(id) {
  if(!/^PUBLIC_CUSTOM_THEME_[a-f0-9-]+$/.test(String(id)) || !customThemes_()[id]) throw new Error('Only saved custom themes can be deleted.');
  const active=getActiveThemeKey()===id;
  PropertiesService.getDocumentProperties().deleteProperty(id);customThemeCache_=null;
  if(active) switchTheme(DEFAULT_THEME_KEY);
  return themeEditorState();
}
function showThemeEditor() {
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createTemplateFromFile('ThemeEditor').evaluate().setWidth(850).setHeight(750),'Theme Editor');
}
