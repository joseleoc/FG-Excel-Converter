import {
  Document,
  Paragraph,
  TextRun,
  Table,
  TableCell,
  TableRow,
  WidthType,
} from "docx";
import * as XLSX from "xlsx";

export interface InterpreterAssignment {
  "Reference Number"?: string;
  Interpreter?: string;
  "Assignment Date"?: string;
  "School Contact"?: string;
  "Assignment Type"?: string;
  Language?: string;
  "Student Name"?: string;
  School?: string;
  "School Address"?: string;
  "Contact Phone"?: string;
  Approved?: string;
}

export const convertExcelToWordDocument = (excelData: ArrayBuffer) => {
  // Read the Excel file
  const workbook = XLSX.read(excelData);
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const jsonData = XLSX.utils.sheet_to_json<InterpreterAssignment>(worksheet);

  // Create Word document
  return createWordDocument(jsonData);
};

const createWordDocument = (data: InterpreterAssignment[]) => {
  const children = data
    .map((row, index) => {
      if (row.Approved === "Cancelled Charge") return null;
      return [
        new Table({
          borders: {
            right: { style: "single", size: 8 },
            top: { style: "single", size: 8 },
            bottom: { style: "single", size: 8 },
            left: { style: "single", size: 8 },
          },
          width: { type: WidthType.AUTO, size: 100 },
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  width: { type: WidthType.AUTO, size: 100 },
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `${index + 1}. Reference Number: ${
                            row["Reference Number"] || ""
                          }`,
                          color: "#5b9bd5",
                          bold: true,
                        }),
                      ],
                      spacing: { after: 200 },
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `Interpreter: ${row.Interpreter || ""}`,
                          bold: true,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `Session Date / Time: ${formatDate(
                            row["Assignment Date"]
                          )}`,
                          bold: true,
                        }),
                      ],
                    }),
                    new Paragraph(`Requestor: ${row["School Contact"] || ""}`),
                    new Paragraph(`Requestor's email: `),
                    new Paragraph(`Type of Service: ${"Interpretation"}`),
                    new Paragraph(
                      `Type of Interpretation: ${getInterpretationType(
                        row["Assignment Type"]
                      )}`
                    ),
                    new Paragraph(
                      `Type of Session: ${getSessionType(
                        row["Assignment Type"]
                      )}`
                    ),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `Language: `,
                          bold: true,
                        }),
                        new TextRun({
                          text: row.Language || "",
                          bold: true,
                          color: "#00b0f0", // Blue color (you can use any hex color)
                        }),
                      ],
                    }),
                    new Paragraph(`Meeting : ${"Interpretation"}`),
                    new Paragraph(`Student Name: ${row["Student Name"] || ""}`),
                    new Paragraph(`Student ID: ${""}`),
                    new Paragraph(`Student Grade: ${""}`),
                    new Paragraph(
                      `Student Campus: ${
                        row.School?.toLowerCase()
                          .split("school")[0]
                          .toUpperCase() || ""
                      }`
                    ),

                    new Paragraph(`Location: ${row.School || ""}`),
                    new Paragraph(
                      `Meeting Place: ${row["School Address"] || ""}`
                    ),
                    new Paragraph(`Phone #: ${row["Contact Phone"] || ""}`),
                    new Paragraph(`Virtual Meeting Link: ${""}`),
                    new Paragraph({ text: "", spacing: { after: 400 } }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ];
    })
    .filter(Boolean)
    .flat();

  return new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Calibri",
            size: 22, // 11pt (22 half-points)
          },
          paragraph: {
            spacing: { line: 276 }, // 1.15 line spacing
          },
        },
      },
    },
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: "Interpreter Assignment Report",
            heading: "Heading1",
            spacing: { after: 400 },
          }),
          // @ts-expect-error - Record contents is not an array
          ...children,
        ],
      },
    ],
  });
};

// Helper functions
const formatDate = (dateString?: string) => {
  if (!dateString) return "";
  try {
    const date = new Date(dateString);
    return isNaN(date.getTime()) ? dateString : date.toLocaleString();
  } catch {
    return dateString;
  }
};

const getInterpretationType = (assignmentType?: string) => {
  if (!assignmentType) return "In-Person (Face-to-Face)";
  return assignmentType.includes("Remote")
    ? "Remote"
    : "In-Person (Face-to-Face)";
};

const getSessionType = (assignmentType?: string) => {
  if (!assignmentType) return "";
  if (assignmentType.includes("ARD")) return "ARD";
  if (assignmentType.includes("In-Home"))
    return "In-Home Autism Training - (SPED)";
  return assignmentType;
};
