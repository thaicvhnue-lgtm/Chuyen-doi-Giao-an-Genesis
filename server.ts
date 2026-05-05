import express from 'express';
import { createServer as createViteServer } from 'vite';
import multer from 'multer';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { GoogleGenAI } from '@google/genai';
import * as dotenv from 'dotenv';

dotenv.config();

const upload = multer({ storage: multer.memoryStorage() });
const userProvidedKey = 'AIzaSyBX05amxp-3bQq9brkXN5r0KTeD3_dq6ew';
const API_KEY = process.env.GEMINI_API_KEY && process.env.GEMINI_API_KEY !== 'MY_GEMINI_API_KEY' 
  ? process.env.GEMINI_API_KEY 
  : userProvidedKey;

const ai = new GoogleGenAI({ apiKey: API_KEY });

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  app.post('/api/convert', upload.single('file'), async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ error: 'No file uploaded' });
      }

      console.log('Processing uploaded docx...');
      const zip = new AdmZip(req.file.buffer);
      const docXmlEntry = zip.getEntry('word/document.xml');
      if (!docXmlEntry) {
        throw new Error('Invalid Word Document: missing word/document.xml');
      }

      const docXmlStr = docXmlEntry.getData().toString('utf8');
      const doc = new DOMParser().parseFromString(docXmlStr, 'text/xml');
      const body = doc.documentElement.getElementsByTagName('w:body')[0];

      // Extract text helper with MathType/Image detection
      function extractText(node) {
        let text = '';
        if (node.nodeName === 'w:t') {
            text += node.textContent || '';
        } else if (node.nodeName === 'm:oMath' || node.nodeName === 'm:oMathPara') {
            text += ' [CÔNG THỨC TOÁN] ';
        } else if (node.nodeName === 'w:drawing' || node.nodeName === 'v:imagedata' || node.nodeName === 'w:pict') {
            text += ' [HÌNH ẢNH] ';
        } else if (node.nodeName === 'w:object') {
            text += ' [MATHTYPE] ';
        }
        
        if (node.childNodes) {
            for (let i = 0; i < node.childNodes.length; i++) {
                text += extractText(node.childNodes[i]);
            }
        }
        return text;
      }

      const elements = [];
      const nodeMap = {};
      let pIdCounter = 0;
      let tblIdCounter = 0;

      function processNode(node) {
        if (node.nodeName === 'w:p') {
          const id = `p_${pIdCounter++}`;
          const text = extractText(node);
          nodeMap[id] = node.cloneNode(true);
          return { id, type: 'p', text };
        }
        if (node.nodeName === 'w:tbl') {
          const id = `tbl_${tblIdCounter++}`;
          nodeMap[id] = node.cloneNode(true);
          
          const rows = [];
          const trNodes = Array.from(node.childNodes).filter(n => n.nodeName === 'w:tr');
          for (const tr of trNodes) {
             const rowCells = [];
             const tcNodes = Array.from(tr.childNodes).filter(n => n.nodeName === 'w:tc');
             for (let i = 0; i < tcNodes.length; i++) {
                const tc = tcNodes[i];
                const contents = [];
                const tcChildren = Array.from(tc.childNodes);
                for (const child of tcChildren) {
                   if (child.nodeName === 'w:p' || child.nodeName === 'w:tbl') {
                      const processed = processNode(child);
                      if (processed) contents.push(processed);
                   }
                }
                rowCells.push({ cellIndex: i, contents });
             }
             rows.push(rowCells);
          }
          return { id, type: 'tbl', rows };
        }
        return null;
      }

      // Extract paragraphs & tables
      const childNodes = Array.from(body.childNodes);
      for (let i = 0; i < childNodes.length; i++) {
        const processed = processNode(childNodes[i]);
        if (processed) elements.push(processed);
      }

      console.log(`Extracted top-level elements. Sending to Gemini...`);

      // Avoid sending massive arrays if doc is surprisingly huge, but for typical lesson plan it's fine.
      const prompt = `You are a helpful assistant that analyzes a Vietnamese lesson plan document (formatted as "Công văn 5512") and maps its paragraphs/tables to a new structured format.

Input:
A JSON array of elements from the document. Elements can be paragraphs ('p') or tables ('tbl'). Tables contain rows and cells, which in turn contain paragraphs.

Output requirements:
Return ONLY a strictly valid JSON object without markdown formatting (no \`\`\`json blocks).
Structure the JSON exactly like this:
{
  "header": {
    "truong": "Tên trường (leave empty if not found)",
    "to": "Tên tổ",
    "mon": "Tên môn học",
    "gv": "Tên giáo viên",
    "lop": "Tên lớp",
    "tenBai": "Tên bài dạy",
    "tietHoc": "Tiết học / thời gian",
    "ngayDay": "Ngày dạy"
  },
  "A_MucTieu": ["p_1", "p_2"], // Array of element IDs that belong to Mục tiêu
  "B_HocLieu": ["p_3"], // Array of IDs for Thiết bị dạy học / Học liệu
  "C_TienTrinh": [             // Array of activities found in Tiến trình dạy học
    {
      "tenHoatDong": "Tên hoạt động (e.g., Hoạt động 1: Mở đầu/Khởi động)",
      "thoiLuong": "Thời lượng",
      "hoatDongGV_HS": ["p_11", "p_12", "tbl_1"], 
        // CHÚ Ý RẤT QUAN TRỌNG: CHỈ BẢO GỒM CÁC ĐOẠN VĂN (p_...) THUỘC 4 BƯỚC (Chuyển giao, Thực hiện, Báo cáo, Kết luận). 
        // NẾU CHÚNG NẰM TRONG 1 BẢNG (ví dụ Bảng có cột "Hoạt động GV-HS" và cột "Sản phẩm"), BẠN PHẢI CHỌN CÁC ID CỦA ĐOẠN VĂN (p_...) TRONG CỘT "Hoạt động" VÀ TUYỆT ĐỐI BỎ QUA các ID đoạn văn của cột "Sản Phẩm". 
        // TUYỆT ĐỐI KHÔNG chọn ID của toàn bộ bảng (tbl_...) nếu đó là bảng chia bố cục. CHỈ dùng ID toàn bộ bảng (tbl_...) cho các bảng dữ liệu bài tập chuyên môn.
        // BỎ QUA phần Mục tiêu, Nội dung, Sản phẩm học tập và chữ "Tổ chức thực hiện".
      "danhGia": [],
      "caNhanHoa": []
    }
  ],
  "baiTapVeNha": [],
  "rutKinhNghiem": []
}

Input elements:
${JSON.stringify(elements)}
`;

      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: prompt,
        config: {
            temperature: 0.1,
            responseMimeType: 'application/json',
        }
      });

      let jsonStr = response.text || '{}';
      // Clean up in case it still returned markdown
      jsonStr = jsonStr.replace(/^```json/m, '').replace(/```$/m, '').trim();
      const planStructure = JSON.parse(jsonStr);

      console.log('Gemini mapping received. Constructing new doc...');

      // --- Reconstruction Phase ---
      
      // Set landscape orientation
      let sectPr = body.getElementsByTagName('w:sectPr')[0];
      if (sectPr) {
        let pgSzList = sectPr.getElementsByTagName('w:pgSz');
        let pgSz = pgSzList.length > 0 ? pgSzList[0] : null;
        if (!pgSz) {
            pgSz = doc.createElement('w:pgSz');
            sectPr.appendChild(pgSz);
        }
        pgSz.setAttribute('w:w', '16838');
        pgSz.setAttribute('w:h', '11906');
        pgSz.setAttribute('w:orient', 'landscape');
      }

      // Detach all nodes from body except sectPr
      childNodes.forEach(child => {
        if (child.nodeName !== 'w:sectPr') {
          body.removeChild(child);
        }
      });

      // Helper functions for Word XML
      function createTextNode(text, bold = false) {
        const p = doc.createElement('w:p');
        const lines = (text || '').split('\n');
        lines.forEach((line, index) => {
            const r = doc.createElement('w:r');
            if (bold) {
                const rPr = doc.createElement('w:rPr');
                rPr.appendChild(doc.createElement('w:b'));
                r.appendChild(rPr);
            }
            if (index > 0) {
                r.appendChild(doc.createElement('w:br'));
            }
            const t = doc.createElement('w:t');
            t.appendChild(doc.createTextNode(line));
            r.appendChild(t);
            p.appendChild(r);
        });
        return p;
      }

      function buildHeaderTable(headerData) {
        const tbl = doc.createElement('w:tbl');
        const tblPr = doc.createElement('w:tblPr');
        const tblW = doc.createElement('w:tblW');
        tblW.setAttribute('w:w', '5000');
        tblW.setAttribute('w:type', 'pct');
        tblPr.appendChild(tblW);
        const borders = doc.createElement('w:tblBorders');
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(dir => {
            const border = doc.createElement('w:' + dir);
            border.setAttribute('w:val', 'single');
            border.setAttribute('w:sz', '4');
            border.setAttribute('w:color', 'auto');
            borders.appendChild(border);
        });
        tblPr.appendChild(borders);
        tbl.appendChild(tblPr);

        const rows = [
            [ `Trường: ${headerData?.truong || ''}`, `Lớp: ${headerData?.lop || ''}` ],
            [ `Tổ: ${headerData?.to || ''}`, `Tên bài: ${headerData?.tenBai || ''}` ],
            [ `Môn học: ${headerData?.mon || ''}`, `Tiết học: ${headerData?.tietHoc || ''}` ],
            [ `Giáo viên: ${headerData?.gv || ''}`, `Ngày dạy: ${headerData?.ngayDay || ''}` ],
        ];

        rows.forEach(r => {
            const tr = doc.createElement('w:tr');
            r.forEach(label => {
                const tc = doc.createElement('w:tc');
                const tcPr = doc.createElement('w:tcPr');
                const tcW = doc.createElement('w:tcW');
                tcW.setAttribute('w:w', '2500');
                tcW.setAttribute('w:type', 'pct');
                tcPr.appendChild(tcW);
                tc.appendChild(tcPr);
                tc.appendChild(createTextNode(label, true));
                tr.appendChild(tc);
            });
            tbl.appendChild(tr);
        });
        return tbl;
      }

      function buildMainTable(activities) {
        const tbl = doc.createElement('w:tbl');
        const tblPr = doc.createElement('w:tblPr');
        const tblW = doc.createElement('w:tblW');
        tblW.setAttribute('w:w', '5000');
        tblW.setAttribute('w:type', 'pct');
        tblPr.appendChild(tblW);
        const borders = doc.createElement('w:tblBorders');
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(dir => {
            const border = doc.createElement('w:' + dir);
            border.setAttribute('w:val', 'single');
            border.setAttribute('w:sz', '4');
            border.setAttribute('w:color', '000000');
            borders.appendChild(border);
        });
        tblPr.appendChild(borders);
        tbl.appendChild(tblPr);

        const trHead = doc.createElement('w:tr');
        const headers = ['Tiến trình', 'Hoạt động của GV và HS\n(Ghi rõ Hoạt động phân hóa nếu có)', 'Đánh giá\n(Hình thức, cách thức, tiêu chí)', 'Tích hợp SDGs\n(nếu có)'];
        const widths = [750, 2750, 750, 750];

        headers.forEach((h, i) => {
            const tc = doc.createElement('w:tc');
            const tcPr = doc.createElement('w:tcPr');
            const tcW = doc.createElement('w:tcW');
            tcW.setAttribute('w:w', widths[i].toString());
            tcW.setAttribute('w:type', 'pct');
            tcPr.appendChild(tcW);
            
            const shading = doc.createElement('w:shd');
            shading.setAttribute('w:val', 'clear');
            shading.setAttribute('w:color', 'auto');
            shading.setAttribute('w:fill', 'EFEFEF');
            tcPr.appendChild(shading);

            tc.appendChild(tcPr);
            tc.appendChild(createTextNode(h, true));
            trHead.appendChild(tc);
        });
        tbl.appendChild(trHead);

        if (!activities || !Array.isArray(activities)) return tbl;

        activities.forEach(act => {
            const tr = doc.createElement('w:tr');

            const mkCell = (w, ids) => {
                const tc = doc.createElement('w:tc');
                const tcPr = doc.createElement('w:tcPr');
                const tcW = doc.createElement('w:tcW');
                tcW.setAttribute('w:w', w.toString());
                tcW.setAttribute('w:type', 'pct');
                tcPr.appendChild(tcW);
                tc.appendChild(tcPr);

                let added = false;
                if (ids && Array.isArray(ids)) {
                    ids.forEach(id => {
                    if(nodeMap[id]) {
                        tc.appendChild(nodeMap[id]);
                        added = true;
                    }
                    });
                }
                return { tc, added };
            };

            // Tiến trình (Column 1)
            const cellTienTrinh = doc.createElement('w:tc');
            const tcPrTienTrinh = doc.createElement('w:tcPr');
            const tcWTienTrinh = doc.createElement('w:tcW');
            tcWTienTrinh.setAttribute('w:w', widths[0].toString());
            tcWTienTrinh.setAttribute('w:type', 'pct');
            tcPrTienTrinh.appendChild(tcWTienTrinh);
            cellTienTrinh.appendChild(tcPrTienTrinh);
            
            const pTienTrinh = doc.createElement('w:p');
            if (act.tenHoatDong) {
                const r1 = doc.createElement('w:r');
                const rPr1 = doc.createElement('w:rPr');
                rPr1.appendChild(doc.createElement('w:b'));
                r1.appendChild(rPr1);
                const t1 = doc.createElement('w:t');
                t1.appendChild(doc.createTextNode(act.tenHoatDong));
                r1.appendChild(t1);
                pTienTrinh.appendChild(r1);
            }
            if (act.thoiLuong) {
                const pThoiLuong = doc.createElement('w:p');
                const r2 = doc.createElement('w:r');
                const rPr2 = doc.createElement('w:rPr');
                rPr2.appendChild(doc.createElement('w:b'));
                r2.appendChild(rPr2);
                const t2 = doc.createElement('w:t');
                t2.appendChild(doc.createTextNode(act.thoiLuong));
                r2.appendChild(t2);
                pThoiLuong.appendChild(r2);
                cellTienTrinh.appendChild(pTienTrinh);
                cellTienTrinh.appendChild(pThoiLuong);
            } else {
                cellTienTrinh.appendChild(pTienTrinh);
            }

            // Hoạt động GV & HS (Column 2) - only the 4 steps
            const cellMoTa = mkCell(widths[1], act.hoatDongGV_HS || []);
            if (!cellMoTa.added) {
                cellMoTa.tc.appendChild(doc.createElement('w:p'));
            }

            // Đánh giá (Column 3)
            const cellDanhGia = mkCell(widths[2], act.danhGia);
            if (!cellDanhGia.added) cellDanhGia.tc.appendChild(doc.createElement('w:p'));

            // Cá nhân hóa / SDGs (Column 4)
            const cellCaNhan = mkCell(widths[3], act.caNhanHoa);
            if (!cellCaNhan.added) cellCaNhan.tc.appendChild(doc.createElement('w:p'));

            tr.appendChild(cellTienTrinh);
            tr.appendChild(cellMoTa.tc);
            tr.appendChild(cellDanhGia.tc);
            tr.appendChild(cellCaNhan.tc);

            tbl.appendChild(tr);
        });

        return tbl;
      }

      function buildSingleColumnTable(title, ids) {
        const tbl = doc.createElement('w:tbl');
        const tblPr = doc.createElement('w:tblPr');
        const tblW = doc.createElement('w:tblW');
        tblW.setAttribute('w:w', '5000');
        tblW.setAttribute('w:type', 'pct');
        tblPr.appendChild(tblW);
        const borders = doc.createElement('w:tblBorders');
        ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(dir => {
            const border = doc.createElement('w:' + dir);
            border.setAttribute('w:val', 'single');
            border.setAttribute('w:sz', '4');
            border.setAttribute('w:color', '000000');
            borders.appendChild(border);
        });
        tblPr.appendChild(borders);
        tbl.appendChild(tblPr);

        // Title row
        const trHead = doc.createElement('w:tr');
        const tcHead = doc.createElement('w:tc');
        const tcPrHead = doc.createElement('w:tcPr');
        tcHead.appendChild(tcPrHead);
        tcHead.appendChild(createTextNode(title, true));
        trHead.appendChild(tcHead);
        tbl.appendChild(trHead);

        // Content row
        if (ids) {
            const trContent = doc.createElement('w:tr');
            const tcContent = doc.createElement('w:tc');
            const tcPrContent = doc.createElement('w:tcPr');
            tcContent.appendChild(tcPrContent);

            let added = false;
            if (Array.isArray(ids)) {
                ids.forEach(id => {
                    if (nodeMap[id]) {
                        tcContent.appendChild(nodeMap[id]);
                        added = true;
                    }
                });
            }
            if (!added) tcContent.appendChild(doc.createElement('w:p'));
            trContent.appendChild(tcContent);
            tbl.appendChild(trContent);
        }

        return tbl;
      }

      function appendNodes(ids) {
        if (!ids || !Array.isArray(ids)) return;
        ids.forEach(id => {
            if (nodeMap[id]) {
                body.insertBefore(nodeMap[id], sectPr);
            }
        });
      }

      // Appending all sections in order
      body.insertBefore(createTextNode('KẾ HOẠCH BÀI DẠY', true), sectPr);
      body.insertBefore(buildHeaderTable(planStructure.header), sectPr);
      
      body.insertBefore(doc.createElement('w:p'), sectPr); // spacing

      body.insertBefore(createTextNode('A. MỤC TIÊU', true), sectPr);
      appendNodes(planStructure.A_MucTieu);

      body.insertBefore(doc.createElement('w:p'), sectPr); // spacing
      
      body.insertBefore(createTextNode('B. HỌC LIỆU HỌC TẬP', true), sectPr);
      appendNodes(planStructure.B_HocLieu);
      
      body.insertBefore(doc.createElement('w:p'), sectPr); // spacing
      
      body.insertBefore(createTextNode('C. TIẾN TRÌNH BÀI DẠY', true), sectPr);
      body.insertBefore(buildMainTable(planStructure.C_TienTrinh), sectPr);

      body.insertBefore(doc.createElement('w:p'), sectPr); // spacing

      body.insertBefore(createTextNode('Bài tập về nhà (nếu có):', true), sectPr);
      appendNodes(planStructure.baiTapVeNha);

      body.insertBefore(createTextNode('Rút kinh nghiệm/ Suy ngẫm:', true), sectPr);
      appendNodes(planStructure.rutKinhNghiem);

      // Format all texts to Montserrat, size 12pt (24 half-points)
      const runs = doc.getElementsByTagName('w:r');
      for (let i = 0; i < runs.length; i++) {
        const r = runs[i];
        let rPr = r.getElementsByTagName('w:rPr')[0];
        if (!rPr) {
            rPr = doc.createElement('w:rPr');
            r.insertBefore(rPr, r.firstChild);
        }
        
        let rFonts = rPr.getElementsByTagName('w:rFonts')[0];
        if (!rFonts) {
            rFonts = doc.createElement('w:rFonts');
            rPr.appendChild(rFonts);
        }
        rFonts.setAttribute('w:ascii', 'Montserrat');
        rFonts.setAttribute('w:hAnsi', 'Montserrat');
        rFonts.setAttribute('w:cs', 'Montserrat');
        rFonts.removeAttribute('w:asciiTheme');
        rFonts.removeAttribute('w:hAnsiTheme');
        rFonts.removeAttribute('w:cstheme');
        rFonts.removeAttribute('w:eastAsiaTheme');
        
        let sz = rPr.getElementsByTagName('w:sz')[0];
        if (!sz) {
            sz = doc.createElement('w:sz');
            rPr.appendChild(sz);
        }
        sz.setAttribute('w:val', '24');

        let szCs = rPr.getElementsByTagName('w:szCs')[0];
        if (!szCs) {
            szCs = doc.createElement('w:szCs');
            rPr.appendChild(szCs);
        }
        szCs.setAttribute('w:val', '24');
      }

      // Serialize and Save
      const updatedXml = new XMLSerializer().serializeToString(doc);
      zip.updateFile('word/document.xml', Buffer.from(updatedXml, 'utf8'));

      const outBuffer = zip.toBuffer();
      res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      res.set('Content-Disposition', 'attachment; filename="Ke_Hoach_Bai_Day.docx"');
      res.send(outBuffer);

    } catch (err: any) {
      console.error(err);
      let errMsg = err.message || 'Internal Server Error';
      if (errMsg.includes('API key not valid')) {
        errMsg = 'API KEY CHƯA ĐÚNG: Bạn hãy vào Cài đặt (biểu tượng \u2699\ufe0f / Secrets), nhấn trực tiếp vào chữ "MY_GEMINI_API_KEY" ở dòng GEMINI_API_KEY để xóa chữ đó đi và dán mã API Key thật của bạn vào.';
      }
      return res.status(500).json({ error: errMsg });
    }
  });


  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const path = await import('path');
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();

