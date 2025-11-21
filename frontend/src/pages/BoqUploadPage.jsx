// src/pages/BoqUploadPage.jsx
import React, { useState, useRef, useCallback } from 'react';
import './BoqUploadPage.css'; // styles below

// const API_URL = "http://127.0.0.1:8000/upload";
import { BACKEND_URL } from "../config";


export default function BoqUploadPage() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [status, setStatus] = useState('');
  const [sheetUrl, setSheetUrl] = useState('');
  const [sheetName, setSheetName] = useState('');
  const [uploadId, setUploadId] = useState('');
  const inputRef = useRef(null);



  async function handleUpload(file) {
  const formData = new FormData();
  formData.append("file", file);

  const res = await fetch(`${BACKEND_URL}/upload`, {
    method: "POST",
    body: formData,
  });

  if (!res.ok) {
    // this is where "Error: Not Found" was coming from
    const errText = await res.text();
    throw new Error(`Backend error ${res.status}: ${errText}`);
  }

  const data = await res.json();
  console.log("BOQ result:", data);
}
  const resetState = () => {
    setStatus('');
    setSheetUrl('');
    setSheetName('');
    setUploadId('');
  };

  const handleFileSelect = (f) => {
    if (!f) return;
    if (!f.name.toLowerCase().match(/\.(dwg|dxf)$/)) {
      setStatus('❌ Please upload a DWG or DXF file.');
      return;
    }
    setFile(f);
    resetState();
  };

  const onDrop = useCallback((e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
    const droppedFiles = e.dataTransfer.files;
    if (droppedFiles && droppedFiles[0]) {
      handleFileSelect(droppedFiles[0]);
    }
  }, []);

  const onDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (!isDragging) setIsDragging(true);
  };

  const onDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleBrowseClick = () => {
    inputRef.current?.click();
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file) {
      setStatus('Please choose a DWG/DXF file first.');
      return;
    }

    setIsUploading(true);
    setStatus('Uploading and generating BOQ…');
    setSheetUrl('');
    setSheetName('');
    setUploadId('');

    try {
      const formData = new FormData();
      formData.append('file', file);

      

const res = await fetch(`${BACKEND_URL}/process`, {
  method: "POST",
  body: formData,
});


      const data = await res.json();

      if (!res.ok || !data.ok) {
        setStatus(
          '❌ Error: ' +
            (data.detail || data.error || data.message || 'Unknown error')
        );
        setIsUploading(false);
        return;
      }

      setSheetUrl(data.sheetUrl);
      setSheetName(data.sheetName);
      setUploadId(data.uploadId);
      setStatus('✅ BOQ generated successfully.');
    } catch (err) {
      console.error(err);
      setStatus('❌ Network error: ' + err.message);
    } finally {
      setIsUploading(false);
    }
  };

  return (
    <div className="boq-page">
      <div className="boq-card">
        <div className="boq-header">
          <div className="boq-pill">Vizdom · AutoCAD BOQ</div>
          <h1>BOQ Generator</h1>
          <p>
            Upload a <strong>DWG</strong> or <strong>DXF</strong> file and get
            an auto-generated BOQ in Google Sheets.
          </p>
        </div>

        <form className="boq-form" onSubmit={handleSubmit}>
          <div
            className={`boq-dropzone ${isDragging ? 'boq-dropzone--drag' : ''}`}
            onDrop={onDrop}
            onDragOver={onDragOver}
            onDragLeave={onDragLeave}
            onDragEnd={onDragLeave}
          >
            <input
              type="file"
              ref={inputRef}
              style={{ display: 'none' }}
              accept=".dwg,.dxf"
              onChange={(e) => handleFileSelect(e.target.files[0])}
            />

            <div className="boq-dropzone-icon">📐</div>

            {file ? (
              <>
                <div className="boq-file-name">{file.name}</div>
                <div className="boq-file-meta">
                  {(file.size / (1024 * 1024)).toFixed(2)} MB · DWG/DXF
                </div>
                <button
                  type="button"
                  className="boq-secondary-btn"
                  onClick={handleBrowseClick}
                  disabled={isUploading}
                >
                  Change file
                </button>
              </>
            ) : (
              <>
                <p className="boq-dropzone-title">
                  Drag &amp; drop your drawing here
                </p>
                <p className="boq-dropzone-sub">
                  or{' '}
                  <button
                    type="button"
                    className="boq-link-btn"
                    onClick={handleBrowseClick}
                    disabled={isUploading}
                  >
                    browse from your computer
                  </button>
                </p>
                <p className="boq-dropzone-hint">Supported: .dwg, .dxf</p>
              </>
            )}
          </div>

          <div className="boq-actions">
            <button
              type="submit"
              className="boq-primary-btn"
              disabled={!file || isUploading}
            >
              {isUploading ? (
                <span className="boq-loader">
                  <span className="boq-loader-dot" />
                  <span className="boq-loader-dot" />
                  <span className="boq-loader-dot" />
                  Processing…
                </span>
              ) : (
                'Generate BOQ'
              )}
            </button>

            {sheetUrl && (
              <a
                href={sheetUrl}
                target="_blank"
                rel="noreferrer"
                className="boq-ghost-btn"
              >
                Open Google Sheet
              </a>
            )}
          </div>
        </form>

        <div className="boq-footer">
          <div className="boq-status">
            {status && <span>{status}</span>}
            {sheetName && (
              <span className="boq-tag">
                Sheet: <strong>{sheetName}</strong>
              </span>
            )}
            {uploadId && <span className="boq-tag">Run ID: {uploadId}</span>}
          </div>

          <div className="boq-footnote">
            Auto-syncs to your BOQ master sheet. No GCP login required for
            users.
          </div>
        </div>
      </div>
    </div>
  );
}
