import { all } from "axios";
import React, { useState } from "react";
import * as XLSX from "xlsx";
import Accordion from "@mui/material/Accordion";
import AccordionActions from "@mui/material/AccordionActions";
import AccordionSummary from "@mui/material/AccordionSummary";
import AccordionDetails from "@mui/material/AccordionDetails";
import Typography from "@mui/material/Typography";
import ExpandMoreIcon from "@mui/icons-material/ExpandMore";
import Button from "@mui/material/Button";
import Table from "@mui/material/Table";
import TableBody from "@mui/material/TableBody";
import TableCell from "@mui/material/TableCell";
import TableContainer from "@mui/material/TableContainer";
import TableHead from "@mui/material/TableHead";
import TableRow from "@mui/material/TableRow";
import Paper from "@mui/material/Paper";

const fieldMap = {
  5: {
    1: 3,
    5: 8,
    12: 14,
  },
  6: {
    1: 3,
    12: 14,
  },
  7: {
    1: 3,
  },
  9: {
    1: 5,
  },
  10: {
    1: 6,
    13: 14,
  },
  11: {
    1: 5,
  },
  14: {
    1: 2,
    8: 9,
  },
  15: {
    1: 2,
    8: 9,
  },
  16: {
    1: 2,
    8: 9,
  },
  17: {
    1: 2,
    8: 9,
  },
  20: {
    1: 2,
    6: 7,
    13: 14,
  },
  21: {
    1: 2,
    6: 7,
    13: 14,
  },
  22: {
    1: 2,
    6: 7,
    13: 14,
  },
  23: {
    1: 2,
    6: 7,
    13: 14,
  },
  24: {
    1: 2,
    6: 7,
    13: 14,
  },
  25: {
    6: 7,
    13: 14,
  },
  26: {
    6: 7,
    13: 14,
  },
  30: {
    1: 5,
  },
  31: {
    1: 5,
  },
  32: {
    1: 5,
    9: 12,
  },
  33: {
    1: 5,
    9: 12,
  },
  34: {
    1: 5,
    9: 12,
  },
  35: {
    1: 5,
    9: 12,
  },
  36: {
    1: 5,
  },
  37: {
    1: 2,
  },
};

export default function ExcelToJson() {
  const [allFilesData, setAllFilesData] = useState({});
  const [expanded, setExpanded] = useState(null);

  const handleFileUpload = (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const results = {};

    Array.from(files).forEach((file) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        let fileResult = {};

        workbook.SheetNames.forEach((sheetName) => {
          const sheet = workbook.Sheets[sheetName];

          // Convert sheet to JSON, keep all rows and fields
          const json = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
            raw: false, // convert Excel dates to text
            dateNF: "dd-mm-yy", // format dates
          });

          console.log({ json });
          // Keep sheet even if it’s empty
          if (sheetName === "Auftrag") {
            // fileResult[sheetName] = json.map((row, index) => ({
            //   rowNumber: index,
            //   values: row,
            // }));
            // console.log({ fileResult });
            const value = {
              top: {},
              middle: {},
              table: {
                "Zuständiger Sachbearbeiter": {},
                "Betreuer-Daten (AD / Makler)": {},
                "TaskForce-Büro": {},
              },
              bottom: {},
            };

            json.forEach(async (row, index) => {
              if (index >= 5 && index <= 11) {
                const map = fieldMap[index];
                row.forEach((itm, itmIndex) => {
                  if (map?.hasOwnProperty(itmIndex)) {
                    const valueIndex = map[itmIndex];
                    value.top[itm] = row[valueIndex];
                  }
                });
              }
              if (index >= 14 && index <= 17) {
                const map = fieldMap[index];
                row.forEach((itm, itmIndex) => {
                  if (map?.hasOwnProperty(itmIndex)) {
                    const valueIndex = map[itmIndex];
                    value.middle[itm] = row[valueIndex];
                  }
                });
              }
              if (index >= 20 && index <= 26) {
                const map = fieldMap[index];
                // console.log(
                //   "row",
                //   row,
                //   "row length" + " " + row.length,
                //   index,
                //   map
                // );
                // // tableRows.push({ row, map });
                row.forEach((itm, itmIndex) => {
                  if (map?.hasOwnProperty(itmIndex)) {
                    console.log("Table", index, map);
                    const valueIndex = map[itmIndex];

                    const name = itm;
                    const val = row[valueIndex];
                    if (itmIndex === 1) {
                      value.table["Zuständiger Sachbearbeiter"][name] = val;
                    } else if (itmIndex === 6) {
                      value.table["Betreuer-Daten (AD / Makler)"][name] = val;
                    } else if (itmIndex === 13) {
                      value.table["TaskForce-Büro"][name] = val;
                    }
                  }
                });
              }
              if (index >= 30 && index <= 37) {
                const map = fieldMap[index];
                row.forEach((itm, itmIndex) => {
                  if (map?.hasOwnProperty(itmIndex)) {
                    const valueIndex = map[itmIndex];
                    value.bottom[itm] = row[valueIndex];
                  }
                });
              }
            });

            fileResult = value;
            // console.log({ value });
          }
        });

        results[file.name] = fileResult;

        if (Object.keys(results).length === files.length) {
          setAllFilesData(results);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  };
  console.log({ allFilesData });

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold mb-2">Excel to JSON Converter</h2>
      <input
        type="file"
        accept=".xlsx, .xls,.xlsm"
        multiple
        onChange={handleFileUpload}
      />

      {Object.keys(allFilesData).length > 0 && (
        <div className="mt-4 space-y-4">
          {Object.entries(allFilesData).map(([fileName, sheets], index) => (
            <Accordion
              onChange={() => {
                console.log({ fileName, index });
                setExpanded(index);
              }}
              sx={{
                background: expanded === index ? "#e2e2e2" : "white",
              }}
            >
              <AccordionSummary
                expandIcon={<ExpandMoreIcon />}
                aria-controls="panel1-content"
                id="panel1-header"
              >
                <Typography component="span">{fileName}</Typography>
              </AccordionSummary>
              <AccordionDetails>
                {/* {JSON.stringify(allFilesData[fileName])} */}
                <TableContainer component={Paper}>
                  <Table
                    sx={{ minWidth: 650 }}
                    size="small"
                    aria-label="simple table"
                  >
                    <TableHead sx={{ background: "#b5b5b5" }}>
                      <TableRow>
                        {Object.keys(allFilesData[fileName].top).map(
                          (tableCell) => (
                            <TableCell>{tableCell}</TableCell>
                          )
                        )}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {Object.keys(allFilesData[fileName].top).map(
                        (tableCell) => (
                          <TableCell>
                            {allFilesData[fileName].top[tableCell]}
                          </TableCell>
                        )
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
                <h4>Versicherungsnehmer</h4>
                <TableContainer component={Paper}>
                  <Table
                    size="small"
                    sx={{ minWidth: 650 }}
                    aria-label="simple table"
                  >
                    <TableHead sx={{ background: "#b5b5b5" }}>
                      <TableRow>
                        {Object.keys(allFilesData[fileName].middle).map(
                          (tableCell) => (
                            <TableCell>{tableCell}</TableCell>
                          )
                        )}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {Object.keys(allFilesData[fileName].middle).map(
                        (tableCell) => (
                          <TableCell>
                            {allFilesData[fileName].middle[tableCell]}
                          </TableCell>
                        )
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
                <h4>Table</h4>
                <TableContainer component={Paper}>
                  <Table
                    size="small"
                    sx={{ minWidth: 650 }}
                    aria-label="simple table"
                  >
                    <TableHead sx={{ background: "#b5b5b5" }}>
                      <TableRow>
                        {Object.keys(allFilesData[fileName].table).map(
                          (tableCell) => (
                            <TableCell>{tableCell}</TableCell>
                          )
                        )}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {Object.keys(allFilesData[fileName].table).map(
                        (tableCell) => (
                          <TableCell>
                            {(() => {
                              const obj =
                                allFilesData[fileName].table[tableCell];
                              // return JSON.stringify(obj);
                              return (
                                <>
                                  {Object.keys(obj).map((key) => {
                                    return (
                                      <div>
                                        {key}
                                        {obj[key]}
                                      </div>
                                    );
                                  })}
                                </>
                              );
                            })()}
                          </TableCell>
                        )
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
                <h4>Vertragsdaten</h4>

                <TableContainer component={Paper}>
                  <Table
                    size="small"
                    sx={{ minWidth: 650 }}
                    aria-label="simple table"
                  >
                    <TableHead sx={{ background: "#b5b5b5" }}>
                      <TableRow>
                        {Object.keys(allFilesData[fileName].bottom).map(
                          (tableCell) => (
                            <TableCell>{tableCell}</TableCell>
                          )
                        )}
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {Object.keys(allFilesData[fileName].bottom).map(
                        (tableCell) => (
                          <TableCell>
                            {allFilesData[fileName].bottom[tableCell]}
                          </TableCell>
                        )
                      )}
                    </TableBody>
                  </Table>
                </TableContainer>
              </AccordionDetails>
            </Accordion>
          ))}
        </div>
      )}
    </div>
  );
}
