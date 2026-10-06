var brandingCache_;
function brandingDefaults_() {
  return {"name":"HERD TRACKER","signature":"","logoUrl":"data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxMjgiIGhlaWdodD0iMTI4IiB2aWV3Qm94PSIwIDAgMTI4IDEyOCI+PHJlY3QgeD0iNCIgeT0iNCIgd2lkdGg9IjEyMCIgaGVpZ2h0PSIxMjAiIHJ4PSIyOCIgZmlsbD0iIzFBNEE2RSIvPjxwYXRoIGQ9Ik0zOSAyOUMyNCA0NyAyNCA3NyAzOCA5MmMxNCAxNSAzOCAxNSA1MiAwIDE0LTE1IDE0LTQ1LTEtNjNMNzQgNDBjMTAgMTMgMTAgMzIgMyA0MS03IDktMTkgOS0yNiAwLTctOS03LTI4IDMtNDFaIiBmaWxsPSJub25lIiBzdHJva2U9IiNFMkQyQjYiIHN0cm9rZS13aWR0aD0iOCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCIvPjxnIGZpbGw9IiNFMkQyQjYiPjxjaXJjbGUgY3g9IjM1IiBjeT0iNTgiIHI9IjMiLz48Y2lyY2xlIGN4PSIzNyIgY3k9Ijc4IiByPSIzIi8+PGNpcmNsZSBjeD0iOTMiIGN5PSI1OCIgcj0iMyIvPjxjaXJjbGUgY3g9IjkxIiBjeT0iNzgiIHI9IjMiLz48L2c+PC9zdmc+","subtitle":""};
}
function validateBranding_(input) {
  if (!input || typeof input !== 'object') throw new Error('Invalid branding settings.');
  const result={};
  for (const [key,limit] of Object.entries({name:60,signature:60,subtitle:60,logoUrl:2048})) {
    if (typeof input[key] !== 'string') throw new Error('Please complete all branding fields.');
    result[key]=input[key].trim();
    if (result[key].length>limit) throw new Error(key+' is too long (maximum '+limit+' characters).');
  }
  if (!result.name) throw new Error('Please enter a stud name.');
  if (!result.logoUrl) result.logoUrl=brandingDefaults_().logoUrl;
  if (result.logoUrl !== brandingDefaults_().logoUrl && !/^https:\/\/[^\s<>"']+$/i.test(result.logoUrl)) throw new Error('Use a direct HTTPS image link.');
  if (/^https:\/\/drive\.google\.com\//i.test(result.logoUrl)) throw new Error('A Google Drive sharing page is not a direct image link. Use a publicly accessible HTTPS image URL.');
  return result;
}
function getBranding_() {
  if (brandingCache_) return brandingCache_;
  try {
    const saved=PropertiesService.getDocumentProperties().getProperty('PUBLIC_HERD_BRANDING');
    brandingCache_=saved ? validateBranding_(JSON.parse(saved)) : brandingDefaults_();
  } catch (error) { brandingCache_=brandingDefaults_(); }
  return brandingCache_;
}
function getBrandingSettings() { return getBranding_(); }
function saveBrandingSettings(input) {
  const settings=validateBranding_(input);
  PropertiesService.getDocumentProperties().setProperty('PUBLIC_HERD_BRANDING',JSON.stringify(settings));
  brandingCache_=settings;
  return settings;
}
function resetBrandingSettings() {
  PropertiesService.getDocumentProperties().deleteProperty('PUBLIC_HERD_BRANDING');
  brandingCache_=brandingDefaults_();
  return brandingCache_;
}
function showBrandingSettings() {
  const html=HtmlService.createTemplateFromFile('BrandingSettings').evaluate().setWidth(650).setHeight(690);
  SpreadsheetApp.getUi().showModalDialog(html,'Branding');
}

function renderPublicLogo_(className) {
  const escape=value=>String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  return '<img src="'+escape(getBranding_().logoUrl)+'" class="'+escape(className)+'" alt="Stud logo">';
}
