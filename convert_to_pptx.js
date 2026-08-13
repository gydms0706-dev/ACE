const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');

const chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const htmlUrl = 'file:///C:/Users/Administrator/Desktop/ACE/presentation_en.html';
const outputPptxPath = 'C:\\Users\\Administrator\\Desktop\\ACE\\presentation_en.pptx';

async function main() {
  console.log('Starting Headless Chrome...');
  const chrome = spawn(chromePath, [
    '--headless=new',
    '--remote-debugging-port=9222',
    '--window-size=1920,1357',
    '--no-sandbox',
    '--disable-gpu',
    '--hide-scrollbars'
  ]);

  await new Promise(r => setTimeout(r, 1500));

  try {
    const targetRes = await fetch('http://127.0.0.1:9222/json/new?' + encodeURIComponent(htmlUrl), { method: 'PUT' });
    const target = await targetRes.json();
    console.log('Target created:', target.id);

    const ws = new WebSocket(target.webSocketDebuggerUrl);

    function send(method, params = {}, id = 1) {
      return new Promise((resolve) => {
        const handler = (evt) => {
          const msg = JSON.parse(evt.data);
          if (msg.id === id) {
            ws.removeEventListener('message', handler);
            resolve(msg.result);
          }
        };
        ws.addEventListener('message', handler);
        ws.send(JSON.stringify({ id, method, params }));
      });
    }

    ws.onopen = async () => {
      console.log('WebSocket connected. Initializing Page...');
      await send('Page.enable', {}, 1);
      await send('DOM.enable', {}, 2);

      // Wait for page rendering
      await new Promise(r => setTimeout(r, 1500));

      const evalRes = await send('Runtime.evaluate', {
        expression: `
          Array.from(document.querySelectorAll('.slide')).map((el, i) => {
            const r = el.getBoundingClientRect();
            return { index: i, x: r.x, y: r.y, width: r.width, height: r.height };
          })
        `,
        returnByValue: true
      }, 3);

      const slides = (evalRes && evalRes.result && evalRes.result.value) ? evalRes.result.value : (evalRes.value || []);
      console.log(`Found ${slides.length} slides.`);

      const capturedImages = [];
      let reqId = 10;

      for (const s of slides) {
        reqId++;
        console.log(`Capturing Slide ${s.index + 1}/${slides.length}...`);
        const shotRes = await send('Page.captureScreenshot', {
          format: 'png',
          clip: {
            x: s.x,
            y: s.y,
            width: s.width,
            height: s.height,
            scale: 2.0
          },
          captureBeyondViewport: true
        }, reqId);

        const imgPath = path.join(__dirname, `slide_${s.index + 1}.png`);
        fs.writeFileSync(imgPath, Buffer.from(shotRes.data, 'base64'));
        console.log(`Saved: slide_${s.index + 1}.png`);
        capturedImages.push(imgPath);
      }

      console.log('Packaging slides into PPTX file...');
      const pptx = new pptxgen();
      pptx.layout = 'LAYOUT_16x9';

      capturedImages.forEach((imgPath, idx) => {
        const slide = pptx.addSlide();
        slide.addImage({
          path: imgPath,
          x: 0,
          y: 0,
          w: '100%',
          h: '100%'
        });

        // Slide 7 (index 6): Embed 12345678.mp4 video natively in PowerPoint!
        if (idx === 6) {
          const videoPath = 'C:\\Users\\Administrator\\Desktop\\ACE\\12345678.mp4';
          if (fs.existsSync(videoPath)) {
            console.log('Embedding 12345678.mp4 into Slide 7...');
            slide.addMedia({
              type: 'video',
              path: videoPath,
              x: 1.8,
              y: 0.95,
              w: 6.4,
              h: 3.6
            });
          }
        }
      });

      await pptx.writeFile({ fileName: outputPptxPath });
      console.log(`SUCCESS! PPTX generated at: ${outputPptxPath}`);

      ws.close();
      chrome.kill();
      process.exit(0);
    };

  } catch (err) {
    console.error('Execution Error:', err);
    chrome.kill();
    process.exit(1);
  }
}

main();
