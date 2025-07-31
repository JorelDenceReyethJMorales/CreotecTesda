import React, { useState } from 'react';

export default function DownloadLinks({ files }) {
  const [fileNames, setFileNames] = useState(
    files.reduce((acc, file, idx) => ({ ...acc, [idx]: file }), {})
  );

  const handleFileNameChange = (idx, newName) => {
    setFileNames(prev => ({ ...prev, [idx]: newName }));
  };

  const handleDownloadAll = () => {
    files.forEach((file, idx) => {
      downloadFile(file, idx);
    });
  };

  // Download wrapper to fetch blob and apply custom filename
  const downloadFile = async (file, idx) => {
    try {
      const res = await fetch(`http://localhost:5000/generate/files/${file}`);
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileNames[idx] || file;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (e) {
      console.error('Download failed', e);
    }
  };
  
  // Generate direct GET URL for ZIP download
  const getZipUrl = () => {
    const qs = files.map(f => `files=${encodeURIComponent(f)}`).join('&');
    return `http://localhost:5000/generate/zip?${qs}`;
  };

  return (
    <div className="container rounded-md">
      <h3 className="section-header">Download Certificates</h3>
      <div className="mb-4">
        <button
          type="button"
          onClick={handleDownloadAll}
          className="bg-green-600 text-white px-4 py-2 rounded-md hover:bg-green-700"
        >
          Download All Individually
        </button>
      </div>
      <div className="flex flex-col gap-1">
        {files.map((file, idx) => (
          <div key={idx} className="certDetails flex gap-2 bg-[#ffffff1a] p-2 mt-2 items-center rounded-md">
            <input
              type="text"
              value={fileNames[idx] || file}
              onChange={(e) => handleFileNameChange(idx, e.target.value)}
              className="flex-1 mainInput h-1rem rounded-md p-1 bg-[#ffffff1a] text-white border border-[#ffffff3a] focus:border-[#a361ef]"
              placeholder="Enter file name"
            />
            <button
              type="button"
              onClick={() => downloadFile(file, idx)}
              className="text-xs font-semibold text-[#613d89] bg-[#a361ef90] px-3 py-1 rounded-md hover:bg-[#a361ef]"
            >
              Download
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}
