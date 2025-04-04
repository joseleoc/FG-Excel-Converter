import "./FileDropzone.css";
import { useCallback, useState } from "react";
import { useDropzone } from "react-dropzone";
import {
  CheckCircleIcon,
  DocumentArrowUpIcon,
  ArrowUpTrayIcon,
} from "@heroicons/react/24/outline";

interface FileWithPreview extends File {
  preview?: string;
}

export default function FileDropzone({
  onFileAccepted,
}: {
  onFileAccepted: (file: File) => void;
}) {
  const [file, setFile] = useState<FileWithPreview | null>(null);

  const onDrop = useCallback(
    (acceptedFiles: File[]) => {
      if (acceptedFiles.length > 0) {
        const acceptedFile = acceptedFiles[0];
        setFile(
          Object.assign(acceptedFile, {
            preview: URL.createObjectURL(acceptedFile),
          })
        );
        onFileAccepted(acceptedFile);
      }
    },
    [onFileAccepted]
  );

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
        ".xlsx",
      ],
      "application/vnd.ms-excel": [".xls"],
    },
    maxFiles: 1,
  });

  return (
    <div className="container">
      <div
        {...getRootProps()}
        className={`dropzone ${isDragActive ? "active" : ""}`}>
        <input {...getInputProps()} />
        {file ? (
          <div className="file-info">
            <div className="file-success">
              <CheckCircleIcon className="success-icon" />
              <p>File ready: {file.name}</p>
            </div>
          </div>
        ) : (
          <div className="dropzone-content">
            {isDragActive ? (
              <>
                <ArrowUpTrayIcon className="upload-icon" />
                <p>Drop the Excel file here</p>
              </>
            ) : (
              <>
                <DocumentArrowUpIcon className="upload-icon" />
                <p>Drag & drop an Excel file here</p>
                <p>or click to select</p>
              </>
            )}
            <em>(Only *.xlsx and *.xls files will be accepted)</em>
          </div>
        )}
      </div>
    </div>
  );
}
