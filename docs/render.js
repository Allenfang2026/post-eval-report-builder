const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
  const dir = __dirname;
  const htmlPath = 'file://' + path.join(dir, '使用手册.html');
  const pdfPath = path.join(dir, '使用手册.pdf');

  const browser = await puppeteer.launch({ headless: 'new', args: ['--no-sandbox'] });
  const page = await browser.newPage();
  await page.goto(htmlPath, { waitUntil: 'networkidle0' });
  // 用 screen 媒体渲染 PDF，绕开打印字体管线丢字问题
  await page.emulateMediaType('screen');
  // 等字体渲染稳定
  await page.evaluateHandle('document.fonts.ready');
  await new Promise(r => setTimeout(r, 800));

  await page.pdf({
    path: pdfPath,
    format: 'A4',
    printBackground: true,
    margin: { top: '0mm', right: '0mm', bottom: '0mm', left: '0mm' },
    preferCSSPageSize: true,
  });

  await browser.close();
  console.log('PDF written:', pdfPath);
})();
