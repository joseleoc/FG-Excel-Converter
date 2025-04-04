import "./ExcelUploadPage.css";
import { useState } from "react";
import { Packer } from "docx";
import { saveAs } from "file-saver";
import { CheckCircleIcon, ArrowUpTrayIcon } from "@heroicons/react/24/outline";
import FileDropzone from "../components/FileDropzone/FileDropzone";
import { convertExcelToWordDocument } from "../functions/excelToWordConverter";

export default function ExcelToWordConverter() {
  const [file, setFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const [conversionSuccess, setConversionSuccess] = useState(false);

  const handleFileAccepted = (file: File) => {
    setFile(file);
    setConversionSuccess(false);
  };

  const convertToWord = async () => {
    if (!file) return;

    setIsConverting(true);
    setConversionSuccess(false);

    try {
      // Read the Excel file
      const data = await file.arrayBuffer();

      // Use the conversion function from the separate file
      const doc = convertExcelToWordDocument(data);

      // Generate and download the Word file
      const blob = await Packer.toBlob(doc);
      saveAs(blob, "Interpreter_Assignments.docx");

      setConversionSuccess(true);
    } catch (error) {
      console.error("Conversion error:", error);
      alert(
        "Error converting file. Please check the file format and try again."
      );
    } finally {
      setIsConverting(false);
    }
  };

  return (
    <div className="container">
      <h1 className="title">Excel to Word Converter</h1>
      <p className="description">
        Upload your Excel file to convert interpreter assignments to a Word
        document
      </p>

      <FileDropzone onFileAccepted={handleFileAccepted} />

      {file && (
        <div className="actions">
          <button
            onClick={convertToWord}
            disabled={isConverting}
            className={`convert-button ${isConverting ? "converting" : ""}`}>
            {isConverting ? (
              "Converting..."
            ) : (
              <>
                <ArrowUpTrayIcon className="button-icon" />
                Convert to Word
              </>
            )}
          </button>

          {conversionSuccess && (
            <div className="success-message">
              <CheckCircleIcon className="success-icon" />
              <span>Conversion successful! File downloaded.</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
