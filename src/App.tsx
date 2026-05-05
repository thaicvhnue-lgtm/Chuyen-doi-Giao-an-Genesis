import { useState } from 'react';
import { UploadCloud, FileType2, Loader2, CheckCircle2, Download, AlertCircle } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [convertSuccess, setConvertSuccess] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setError(null);
      setConvertSuccess(false);
      setDownloadUrl(null);
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const droppedFile = e.dataTransfer.files[0];
      if (droppedFile.name.endsWith('.docx')) {
        setFile(droppedFile);
        setError(null);
        setConvertSuccess(false);
        setDownloadUrl(null);
      } else {
        setError('Vui lòng tải lên file .docx (Word document).');
      }
    }
  };

  const handleConvert = async () => {
    if (!file) return;

    setIsConverting(true);
    setError(null);

    const formData = new FormData();
    formData.append('file', file);

    try {
      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || `Lỗi HTTP: ${response.status}`);
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setConvertSuccess(true);
      
      // Auto-trigger download
      const a = document.createElement('a');
      a.href = url;
      a.download = `KeHoachBaiDay_${file.name}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      
    } catch (err: any) {
      console.error('Lỗi chuyển đổi:', err);
      setError(err.message || 'Đã xảy ra lỗi trong quá trình chuyển đổi.');
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans">
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="bg-blue-600 p-2 rounded-lg text-white">
            <FileType2 className="w-5 h-5" />
          </div>
          <h1 className="text-xl font-semibold text-gray-900 tracking-tight">Chuẩn Hóa Kế Hoạch Bài Dạy</h1>
        </div>
        <div className="text-sm text-gray-500">Hỗ trợ Công văn 5512 &rarr; Mẫu Kế hoạch bài dạy mới (không làm hỏng MathType)</div>
      </header>

      <main className="flex-1 max-w-4xl w-full mx-auto p-6 md:p-12 flex flex-col mt-4">
        <div className="text-center mb-10">
          <h2 className="text-3xl font-bold text-gray-900 mb-4 tracking-tight">
            Chuyển đổi giáo án dễ dàng
          </h2>
          <p className="text-lg text-gray-600 max-w-2xl mx-auto">
            Công cụ sử dụng AI để tự động phân tích cấu trúc văn bản, giữ nguyên 100% công thức Toán (MathType / OMML), hình ảnh và bố cục đoạn văn so với bản gốc.
          </p>
        </div>

        <div className="bg-white p-8 rounded-2xl shadow-sm border border-gray-100 flex flex-col items-center">
          
          <div 
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            className={`w-full max-w-xl p-10 border-2 border-dashed rounded-xl flex flex-col items-center justify-center transition-colors ${
              file ? 'border-blue-300 bg-blue-50/50' : 'border-gray-300 hover:border-gray-400 hover:bg-gray-50'
            }`}
          >
            {file ? (
              <div className="flex flex-col items-center text-center gap-3">
                <div className="w-16 h-16 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center">
                  <FileType2 className="w-8 h-8" />
                </div>
                <div>
                  <h3 className="text-lg font-medium text-gray-900">{file.name}</h3>
                  <p className="text-sm text-gray-500">{(file.size / 1024).toFixed(1)} KB</p>
                </div>
                {!isConverting && !convertSuccess && (
                  <button 
                    onClick={() => setFile(null)}
                    className="text-sm text-red-600 hover:text-red-700 font-medium mt-2"
                  >
                    Hủy / Chọn file khác
                  </button>
                )}
              </div>
            ) : (
              <div className="flex flex-col items-center text-center gap-4">
                <div className="w-16 h-16 bg-gray-100 text-gray-400 rounded-full flex items-center justify-center">
                  <UploadCloud className="w-8 h-8" />
                </div>
                <div>
                  <h3 className="text-lg font-medium text-gray-900">Tải lên giáo án (.docx)</h3>
                  <p className="text-sm text-gray-500 mt-1">Kéo thả file vào đây hoặc nhấn để duyệt</p>
                </div>
                <input
                  type="file"
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  onChange={handleFileChange}
                  className="hidden"
                  id="file-upload"
                />
                <label 
                  htmlFor="file-upload"
                  className="mt-2 text-sm font-medium text-blue-600 hover:text-blue-700 cursor-pointer bg-blue-50 px-4 py-2 rounded-lg transition-colors"
                >
                  Duyệt file trên máy
                </label>
              </div>
            )}
          </div>

          {error && (
            <motion.div 
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className="w-full max-w-xl mt-6 p-4 bg-red-50 text-red-700 rounded-xl border border-red-100 flex items-start gap-3"
            >
              <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
              <div className="text-sm font-medium">{error}</div>
            </motion.div>
          )}

          {file && !convertSuccess && (
            <div className="mt-8 w-full max-w-xl flex justify-center">
              <button
                onClick={handleConvert}
                disabled={isConverting}
                className={`flex items-center gap-2 px-8 py-3 rounded-full font-medium text-white transition-all shadow-sm ${
                  isConverting 
                    ? 'bg-blue-400 cursor-not-allowed' 
                    : 'bg-blue-600 hover:bg-blue-700 hover:shadow-md active:scale-95'
                }`}
              >
                {isConverting ? (
                  <>
                    <Loader2 className="w-5 h-5 animate-spin" />
                    Đang xử lý bằng AI (có thể mất 15-30s)...
                  </>
                ) : (
                  <>
                    <FileType2 className="w-5 h-5" />
                    Bắt đầu chuyển đổi
                  </>
                )}
              </button>
            </div>
          )}

          <AnimatePresence>
            {convertSuccess && downloadUrl && (
              <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                className="mt-8 w-full max-w-xl p-8 bg-green-50 rounded-2xl border border-green-100 flex flex-col items-center text-center"
              >
                <div className="w-16 h-16 bg-green-100 text-green-600 rounded-full flex items-center justify-center mb-4">
                  <CheckCircle2 className="w-8 h-8" />
                </div>
                <h3 className="text-xl font-semibold text-gray-900 mb-2">Chuyển đổi thành công!</h3>
                <p className="text-gray-600 mb-6">Giáo án của bạn đã được cấu trúc lại theo mẫu mới với đầy đủ công thức và hình ảnh.</p>
                
                <div className="flex gap-4">
                  <a
                    href={downloadUrl}
                    download={`KeHoachBaiDay_${file.name}`}
                    className="flex items-center gap-2 px-6 py-2.5 bg-white text-gray-800 border border-gray-300 font-medium rounded-xl hover:bg-gray-50 transition-colors shadow-sm"
                  >
                    <Download className="w-5 h-5" />
                    Tải xuống lại
                  </a>
                  <button
                    onClick={() => {
                      setFile(null);
                      setConvertSuccess(false);
                      setDownloadUrl(null);
                    }}
                    className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 text-white font-medium rounded-xl hover:bg-blue-700 transition-colors shadow-sm"
                  >
                    Chuyển đổi file khác
                  </button>
                </div>
              </motion.div>
            )}
          </AnimatePresence>

        </div>
        
        <div className="mt-12 w-full max-w-3xl mx-auto grid grid-cols-1 md:grid-cols-3 gap-6 text-center text-sm text-gray-600">
            <div>
              <div className="font-semibold text-gray-900 mb-1">An toàn tuyệt đối</div>
              File của bạn được xử lý trực tiếp trên server và không được lưu trữ dài hạn.
            </div>
            <div>
              <div className="font-semibold text-gray-900 mb-1">Giữ nguyên công thức</div>
              Bằng kỹ thuật bóc tách và chèn XML tinh vi, mọi công thức MathType sẽ không bao giờ bị ảnh hưởng.
            </div>
            <div>
              <div className="font-semibold text-gray-900 mb-1">Hỗ trợ AI phân tích</div>
              Cấu trúc bản thảo sẽ được AI tự động nhận dạng và điền vào các bảng phù hợp.
            </div>
        </div>
      </main>
    </div>
  );
}
