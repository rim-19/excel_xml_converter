import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Checkbox } from "@/components/ui/checkbox";
import { Progress } from "@/components/ui/progress";
import { useToast } from "@/hooks/use-toast";
import { Upload, FileSpreadsheet, Download, CheckCircle2, AlertCircle } from "lucide-react";

interface SheetData {
  name: string;
  data: any[][];
}

export const ExcelToXmlConverter = () => {
  const [file, setFile] = useState<File | null>(null);
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [selectedSheets, setSelectedSheets] = useState<Set<string>>(new Set());
  const [xmlOutput, setXmlOutput] = useState<string>("");
  const [processing, setProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [dragActive, setDragActive] = useState(false);
  const { toast } = useToast();

  // Handle file upload
  const handleFileUpload = useCallback(async (uploadedFile: File) => {
    if (!uploadedFile) return;

    const validTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
      "text/csv"
    ];

    if (!validTypes.includes(uploadedFile.type) && !uploadedFile.name.match(/\.(xlsx|xls|csv)$/i)) {
      toast({
        title: "Invalid File Type",
        description: "Please upload a valid Excel file (.xlsx, .xls, or .csv)",
        variant: "destructive",
      });
      return;
    }

    setFile(uploadedFile);
    setProcessing(true);
    setProgress(20);

    try {
      const arrayBuffer = await uploadedFile.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: "array" });
      
      setProgress(50);

      const extractedSheets: SheetData[] = workbook.SheetNames.map(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        return { name: sheetName, data: data as any[][] };
      });

      setSheets(extractedSheets);
      setSelectedSheets(new Set(extractedSheets.map(s => s.name)));
      setProgress(100);
      
      toast({
        title: "File Loaded Successfully",
        description: `${extractedSheets.length} sheet(s) found`,
      });
    } catch (error) {
      toast({
        title: "Error Reading File",
        description: "Unable to parse the Excel file. Please check the file format.",
        variant: "destructive",
      });
      setFile(null);
    } finally {
      setProcessing(false);
    }
  }, [toast]);

  // Drag and drop handlers
  const handleDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  };

  // Toggle sheet selection
  const toggleSheet = (sheetName: string) => {
    const newSelected = new Set(selectedSheets);
    if (newSelected.has(sheetName)) {
      newSelected.delete(sheetName);
    } else {
      newSelected.add(sheetName);
    }
    setSelectedSheets(newSelected);
  };

  // Convert to XML
  const convertToXml = useCallback(() => {
    if (sheets.length === 0 || selectedSheets.size === 0) {
      toast({
        title: "No Sheets Selected",
        description: "Please select at least one sheet to convert",
        variant: "destructive",
      });
      return;
    }

    setProcessing(true);
    setProgress(0);

    try {
      let xml = '<?xml version="1.0" encoding="UTF-8"?>\n<Workbook>\n';
      
      const selectedSheetsArray = sheets.filter(s => selectedSheets.has(s.name));
      const progressStep = 100 / selectedSheetsArray.length;

      selectedSheetsArray.forEach((sheet, sheetIndex) => {
        xml += `  <Sheet name="${escapeXml(sheet.name)}">\n`;
        
        if (sheet.data.length > 0) {
          const headers = sheet.data[0];
          
          // Process rows (skip header row)
          for (let i = 1; i < sheet.data.length; i++) {
            const row = sheet.data[i];
            if (row.some(cell => cell !== "")) {
              xml += '    <Row>\n';
              
              row.forEach((cell, cellIndex) => {
                const tagName = headers[cellIndex] 
                  ? sanitizeTagName(String(headers[cellIndex]))
                  : `Column${cellIndex + 1}`;
                const value = escapeXml(String(cell));
                xml += `      <${tagName}>${value}</${tagName}>\n`;
              });
              
              xml += '    </Row>\n';
            }
          }
        }
        
        xml += '  </Sheet>\n';
        setProgress((sheetIndex + 1) * progressStep);
      });

      xml += '</Workbook>';
      setXmlOutput(xml);
      
      toast({
        title: "Conversion Complete",
        description: `Successfully converted ${selectedSheets.size} sheet(s) to XML`,
      });
    } catch (error) {
      toast({
        title: "Conversion Error",
        description: "An error occurred during conversion. Please try again.",
        variant: "destructive",
      });
    } finally {
      setProcessing(false);
    }
  }, [sheets, selectedSheets, toast]);

  // Download XML
  const downloadXml = () => {
    if (!xmlOutput) return;

    const blob = new Blob([xmlOutput], { type: "application/xml" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = file ? `${file.name.replace(/\.[^/.]+$/, "")}.xml` : "output.xml";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    toast({
      title: "Download Complete",
      description: "XML file has been saved successfully",
    });
  };

  // Helper functions
  const escapeXml = (str: string): string => {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  };

  const sanitizeTagName = (tag: string): string => {
    return tag
      .replace(/[^a-zA-Z0-9_-]/g, "_")
      .replace(/^[^a-zA-Z_]/, "_");
  };

  return (
    <div className="min-h-screen bg-background p-6">
      <div className="max-w-5xl mx-auto space-y-6">
        {/* Header */}
        <div className="text-center space-y-2">
          <h1 className="text-3xl font-semibold text-foreground">Excel to XML Converter</h1>
          <p className="text-muted-foreground">Convert Excel files to structured XML format</p>
        </div>

        {/* Upload Section */}
        <Card className="p-8">
          <div
            className={`border-2 border-dashed rounded-lg p-12 text-center transition-colors ${
              dragActive ? "border-accent bg-accent/5" : "border-border"
            }`}
            onDragEnter={handleDrag}
            onDragLeave={handleDrag}
            onDragOver={handleDrag}
            onDrop={handleDrop}
          >
            {!file ? (
              <div className="space-y-4">
                <Upload className="w-12 h-12 mx-auto text-muted-foreground" />
                <div>
                  <p className="text-foreground font-medium mb-2">
                    Drag and drop your Excel file here
                  </p>
                  <p className="text-sm text-muted-foreground mb-4">
                    or click to browse (.xlsx, .xls, .csv)
                  </p>
                  <input
                    type="file"
                    id="file-upload"
                    accept=".xlsx,.xls,.csv"
                    onChange={(e) => e.target.files && handleFileUpload(e.target.files[0])}
                    className="hidden"
                  />
                  <label htmlFor="file-upload">
                    <Button asChild>
                      <span>Select Excel File</span>
                    </Button>
                  </label>
                </div>
              </div>
            ) : (
              <div className="space-y-4">
                <FileSpreadsheet className="w-12 h-12 mx-auto text-accent" />
                <div>
                  <p className="text-foreground font-medium">{file.name}</p>
                  <p className="text-sm text-muted-foreground">
                    {(file.size / 1024).toFixed(2)} KB
                  </p>
                </div>
                <Button variant="secondary" onClick={() => {
                  setFile(null);
                  setSheets([]);
                  setXmlOutput("");
                }}>
                  Change File
                </Button>
              </div>
            )}
          </div>

          {processing && progress < 100 && (
            <div className="mt-6 space-y-2">
              <Progress value={progress} />
              <p className="text-sm text-center text-muted-foreground">Loading file...</p>
            </div>
          )}
        </Card>

        {/* Sheet Selection */}
        {sheets.length > 0 && !xmlOutput && (
          <Card className="p-6">
            <h2 className="text-xl font-semibold mb-4 text-foreground">Select Sheets to Convert</h2>
            <div className="space-y-3 mb-6">
              {sheets.map((sheet) => (
                <div key={sheet.name} className="flex items-center space-x-3 p-3 rounded border border-border hover:bg-muted/50 transition-colors">
                  <Checkbox
                    id={sheet.name}
                    checked={selectedSheets.has(sheet.name)}
                    onCheckedChange={() => toggleSheet(sheet.name)}
                  />
                  <label
                    htmlFor={sheet.name}
                    className="flex-1 cursor-pointer font-medium text-foreground"
                  >
                    {sheet.name}
                    <span className="ml-2 text-sm text-muted-foreground">
                      ({sheet.data.length - 1} rows)
                    </span>
                  </label>
                </div>
              ))}
            </div>
            <Button
              onClick={convertToXml}
              disabled={selectedSheets.size === 0 || processing}
              className="w-full"
            >
              {processing ? "Converting..." : "Convert to XML"}
            </Button>
            {processing && (
              <div className="mt-4">
                <Progress value={progress} />
              </div>
            )}
          </Card>
        )}

        {/* XML Output */}
        {xmlOutput && (
          <Card className="p-6">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                <CheckCircle2 className="w-5 h-5 text-success" />
                <h2 className="text-xl font-semibold text-foreground">Conversion Complete</h2>
              </div>
              <Button onClick={downloadXml} className="gap-2">
                <Download className="w-4 h-4" />
                Save XML
              </Button>
            </div>
            
            <div className="bg-muted rounded-lg p-4 max-h-96 overflow-auto">
              <pre className="text-sm text-foreground font-mono whitespace-pre-wrap break-all">
                {xmlOutput}
              </pre>
            </div>

            <div className="mt-4 flex gap-3">
              <Button
                variant="secondary"
                onClick={() => {
                  setXmlOutput("");
                  setFile(null);
                  setSheets([]);
                }}
                className="flex-1"
              >
                Convert Another File
              </Button>
            </div>
          </Card>
        )}

        {/* Info Section */}
        {!file && (
          <Card className="p-6 bg-card">
            <div className="flex gap-3">
              <AlertCircle className="w-5 h-5 text-accent flex-shrink-0 mt-0.5" />
              <div className="space-y-2">
                <h3 className="font-medium text-foreground">Output Format</h3>
                <p className="text-sm text-muted-foreground">
                  Each sheet is wrapped in a <code className="bg-muted px-1.5 py-0.5 rounded">&lt;Sheet name="..."&gt;</code> element.
                  Rows are contained within <code className="bg-muted px-1.5 py-0.5 rounded">&lt;Row&gt;</code> tags,
                  with column headers used as XML tag names.
                </p>
              </div>
            </div>
          </Card>
        )}
      </div>
    </div>
  );
};
export default ExcelToXmlConverter;