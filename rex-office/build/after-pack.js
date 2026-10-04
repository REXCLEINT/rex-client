// Setzt Icon und Versionsinfos der Windows-.exe (64-bit rcedit, funktioniert auf Windows und – mit Wine – auf Linux)
const path = require('path');
const rcedit = require('rcedit');

exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== 'win32') return;
  const info = context.packager.appInfo;
  const exe = path.join(context.appOutDir, `${info.productFilename}.exe`);
  await rcedit(exe, {
    icon: path.join(__dirname, 'icon.ico'),
    'file-version': info.version,
    'product-version': info.version,
    'version-string': {
      FileDescription: 'REX Office',
      ProductName: 'REX Office',
      CompanyName: 'REX',
      LegalCopyright: 'Copyright REX',
      InternalName: 'REX Office',
      OriginalFilename: `${info.productFilename}.exe`
    }
  });
};
