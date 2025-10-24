import { all } from "axios";
import React, { useEffect, useState } from "react";
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
import TextField from "@mui/material/TextField";
import Autocomplete from "@mui/material/Autocomplete";
import { Email } from "@mui/icons-material";
import { Box } from "@mui/material";

const ZOHO = window.ZOHO;

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

const per_page = 200;

const useDebouncedValue = (inputValue, delay) => {
  const [loading, setLoading] = useState(false);
  const [debouncedValue, setDebouncedValue] = useState(inputValue);

  useEffect(() => {
    const handler = setTimeout(() => {
      setDebouncedValue(inputValue);
    }, delay);

    return () => {
      clearTimeout(handler);
    };
  }, [inputValue, delay]);

  return debouncedValue;
};

export default function ExcelToJson() {
  const [allFilesData, setAllFilesData] = useState({});
  const [expanded, setExpanded] = useState(null);
  const [searchValue, setSearchValue] = useState(null);
  //Autocomplete
  const [loading, setLoading] = useState(false);
  const [deals, setDeals] = useState([]);
  const [previousSearch, setPreviousSearch] = useState([]);
  const [initailLoading, setInitialLoading] = useState(true);
  const [zohoInitialized, setZohoInitialized] = useState(false);
  const [optionsForRole, setOptionsForRole] = useState(null);
  const [entityInfo, setEntityInfo] = useState(null);
  const [options, setOptions] = useState([]);
  const [value, setValue] = useState(null);
  const [selectedDeals, setSelectedDeals] = useState({});
  const [selectedFiles, setSelectedFiles] = useState([]);

  const debouncedSearchTerm = useDebouncedValue(searchValue, 500);

  const handlePreviousSearch = ({ search }) => {
    let temp = Object.keys(previousSearch)?.filter((pet) => {
      var regexObj = new RegExp("^" + pet, "i");
      if (regexObj.test(search)) return true;
    });

    let result = "";
    if (temp.length != 0) {
      result = temp?.reduce(function (a, b) {
        return a.length > b.length ? a : b;
      });
    }

    return result;
  };

  const handleSearch = async ({ search, page, previousData = [] }) => {
    console.log(
      "handle previous search",
      search,
      !!handlePreviousSearch({ search: search }),
      previousData
    );
    try {
      if (handlePreviousSearch({ search: search }) == "") {
        setLoading((prev) => true);
        try {
          let dealsResp = await ZOHO.CRM.API.searchRecord({
            Entity: "Deals",
            Type: "criteria",
            Query:
              "((Deal_Name:starts_with:" +
              encodeURI("*" + search + "*") +
              ")and(Stage:equals:Auftrag Erstellt))",
            // Query: "(Email:Contains:" + Number(search) + ")",
            per_page: per_page,
            page: page,
            sort_order: "asc",
          });
          console.log({ dealsResp: dealsResp });

          if (!dealsResp?.data) {
            setLoading((prev) => false);
            if (Number(search)) {
              return;
            } else {
              setPreviousSearch((prev) => {
                return {
                  ...prev,
                  [`${search}`]: {
                    data: [...previousData],
                    page: page,
                    more_records: false,
                  },
                };
              });
              setDeals((prev) => [...previousData]);
              return;
            }
          }
          if (dealsResp?.info?.more_records && page < 1) {
            // Call again
            return handleSearch({
              search: search,
              page: page + 1,
              previousData: [...previousData, ...dealsResp?.data],
            });
          } else {
            setPreviousSearch((prev) => {
              return {
                ...prev,
                [`${search}`]: {
                  data: [...previousData, ...dealsResp?.data],
                  page: page,
                  more_records: dealsResp?.info?.more_records,
                },
              };
            });
          }

          setLoading((prev) => false);
          setDeals((prev) => [...previousData, ...dealsResp?.data]);

          return;
        } catch (error) {
          setLoading((prev) => false);
          if (Number(search)) {
            return;
          } else {
            setPreviousSearch((prev) => {
              return {
                ...prev,
                [`${search}`]: {
                  data: [...previousData],
                  page: page,
                  more_records: false,
                },
              };
            });
            setDeals((prev) => [...previousData]);

            return;
          }
        }
      } else if (handlePreviousSearch({ search: search }) == search) {
        //we dont need to do anything here
      } else {
        try {
          let previousResult =
            previousSearch[`${handlePreviousSearch({ search: search })}`];
          if (previousResult?.more_records) {
            setLoading((prev) => true);
            const dealsResp = await ZOHO.CRM.API.searchRecord({
              Entity: "Deals",
              Type: "criteria",
              Query:
                "((Deal_Name:starts_with:" +
                encodeURI("*" + search + "*") +
                ")and(Stage:equals:Auftrag Erstellt))",
              per_page: per_page,
              sort_order: "asc",
              page: page,
            });
            console.log({ dealsResp: dealsResp });
            if (!dealsResp?.data) {
              setLoading((prev) => false);
              if (Number(search)) {
                return;
              } else {
                setPreviousSearch((prev) => {
                  return {
                    ...prev,
                    [`${search}`]: {
                      data: [...previousData],
                      page: page,
                      more_records: false,
                    },
                  };
                });
                setDeals((prev) => [...previousData]);

                return;
              }
            }
            if (dealsResp?.info?.more_records && page < 1) {
              console.log({
                search: search,
                page: page + 1,
                previousData: [...previousData, ...dealsResp?.data],
              });
              // Call again
              return handleSearch({
                search: search,
                page: page + 1,
                previousData: [...previousData, ...dealsResp?.data],
              });
            } else {
              setPreviousSearch((prev) => {
                return {
                  ...prev,
                  [`${search}`]: {
                    data: [...previousData, ...dealsResp?.data],
                    page: page,
                    more_records: dealsResp?.info?.more_records,
                  },
                };
              });
            }

            setLoading((prev) => false);
            setDeals((prev) => [...previousData, ...dealsResp?.data]);

            return;
          }
        } catch (error) {
          setLoading((prev) => false);
          if (Number(search)) {
            return;
          } else {
            setPreviousSearch((prev) => {
              return {
                ...prev,
                [`${search}`]: {
                  data: [...previousData],
                  page: page,
                  more_records: false,
                },
              };
            });
            setDeals((prev) => [...previousData]);

            return;
          }
        }
      }
    } catch (error) {
      setLoading((prev) => false);
      console.log({ error });
    }
  };

  useEffect(() => {
    ZOHO.embeddedApp.on("PageLoad", async function (entityData) {
      console.log("PageLoad", entityData);
      setEntityInfo(entityData);

      // encodeURI("Auftrag Erstellt")
      let dealsResp = await ZOHO.CRM.API.searchRecord({
        Entity: "Deals",
        Type: "criteria",
        Query: "(Stage:equals:Auftrag Erstellt)",
        per_page: 200,
        page: 1,
        sort_order: "asc",
      });
      // console.log("deals fields", dealsResp);
      setDeals(dealsResp.data);

      setInitialLoading(false);
      ZOHO.CRM.UI.Resize({ height: "1200", width: "1500" }).then(function (
        data
      ) {
        // console.log(data);
      });
    });

    ZOHO.embeddedApp.init().then(() => {
      setZohoInitialized(true);
    });
  }, []);

  useEffect(() => {
    // searchValue?.length >= 3 &&
    handleSearch({
      search: debouncedSearchTerm,
      page: 1,
      previousData: deals,
    });
  }, [debouncedSearchTerm]);

  // useEffect(() => {
  //   const data = {};

  //   deals.forEach((option) => {
  //     if (searchValue) {
  //       const isFound = watch("test").find(
  //         (row, rowIndex) =>
  //           row?.PCL_Contact_section?.id !== option?.id &&
  //           (option.Email?.toLowerCase().includes(searchValue?.toLowerCase()) ||
  //             option.Full_Name?.toLowerCase().includes(
  //               searchValue.toLowerCase()
  //             ))
  //       );

  //       if (isFound) data[option.id] = option;
  //     } else {
  //       data[option.id] = option;
  //     }
  //   });

  //   setOptions(Object.values(data));
  // }, [deals]);

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
            fileResult[sheetName] = json.map((row, index) => ({
              rowNumber: index,
              values: row,
            }));

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
              if (index === 38) {
                value.bottom["Termin_Info"] = row[1];
              }
            });

            fileResult = value;
            // console.log({ value });
          }
        });

        results[file.name] = fileResult;
        console.log("file", results[file.name]);

        if (Object.keys(results).length === files.length) {
          setAllFilesData(results);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  };

  function convertToISO(dateStr) {
    const [month, day, year] = dateStr.split("/");

    // Always assume 2-digit years are 2000+
    const fullYear = year.length === 2 ? `20${year.padStart(2, "0")}` : year;

    return `${fullYear}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }

  return (
    <Box sx={{ width: "100%" }}>
      <Box sx={{ width: "96%", mx: "auto" }}>
        <Typography
          sx={{
            fontSize: "1.5rem",
            fontWeight: "bold",
            textAlign: "center",
            mt: 3,
            mb: 2,
          }}
        >
          Excel to Update Deals
        </Typography>

        <Typography
          sx={{
            fontWeight: "bold",
            mb: 1,
          }}
        >
          Upload Excel Files
        </Typography>

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
                  <Typography component="span" sx={{ width: "200px" }}>
                    {fileName}
                  </Typography>
                </AccordionSummary>

                <AccordionDetails>
                  <form
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "10px",
                      marginBottom: "20px",
                    }}
                    onSubmit={async (e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      const files = selectedFiles[fileName] || [];
                      const deal = selectedDeals[fileName];
                      const fileData = allFilesData[fileName];
                      const filesArray = Array.from(files);

                      // console.log({ fileData });
                      const emailForContact = fileData?.middle?.["E-Mail:"];
                      let contactId = "";
                      let contactIdADAgentur = "";
                      if (
                        emailForContact !== "" &&
                        emailForContact !== null &&
                        emailForContact !== undefined
                      ) {
                        const contactSearchResp =
                          await ZOHO.CRM.API.searchRecord({
                            Entity: "Contacts",
                            Type: "email",
                            Query: emailForContact,
                          });
                        console.log(contactSearchResp);
                        contactId = contactSearchResp?.data?.[0]?.id;
                      }

                      let emailForContactADAgentur =
                        fileData?.table?.["Betreuer-Daten (AD / Makler)"]?.[
                          "E-Mail:"
                        ];
                      if (emailForContactADAgentur === emailForContact) {
                        contactIdADAgentur = contactId;
                      } else {
                        if (
                          emailForContactADAgentur !== "" &&
                          emailForContactADAgentur !== null &&
                          emailForContactADAgentur !== undefined
                        ) {
                          const contactADAgenturResp =
                            await ZOHO.CRM.API.searchRecord({
                              Entity: "Contacts",
                              Type: "email",
                              Query: emailForContactADAgentur,
                            });
                          contactIdADAgentur =
                            contactADAgenturResp?.data?.[0]?.id;
                        }
                      }

                      let emailForContactSanierungspartner =
                        fileData?.top?.["Emailadresse:"];

                      let contactIdSanierungspartner = "";

                      if (
                        emailForContactSanierungspartner === emailForContact
                      ) {
                        contactIdSanierungspartner = contactId;
                      } else if (
                        emailForContactSanierungspartner ===
                        emailForContactADAgentur
                      ) {
                        contactIdSanierungspartner = contactIdADAgentur;
                      } else {
                        if (
                          emailForContactSanierungspartner !== "" &&
                          emailForContactSanierungspartner !== null &&
                          emailForContactSanierungspartner !== undefined
                        ) {
                          const contactSanierungspartnerResp =
                            await ZOHO.CRM.API.searchRecord({
                              Entity: "Contacts",
                              Type: "email",
                              Query: emailForContactADAgentur,
                            });
                          contactIdSanierungspartner =
                            contactSanierungspartnerResp?.data?.[0]?.id;
                        }
                      }

                      var config = {
                        Entity: "Deals",
                        APIData: {
                          id: deal?.id,
                          Versicherung: "Sparkassenversicherung",
                          Schadentag: convertToISO(
                            fileData?.top?.["Schadentag: "]
                          ),
                          Schadennummer: fileData?.top?.["Schadennummer:"],
                          Schadenort: fileData?.top?.["Schadenort:"],
                          Termin_Info: fileData?.bottom?.Termin_Info,
                          Schadenart: "GB Leitungswasser",
                          Versicherungsnehmer: contactId
                            ? { id: contactId }
                            : {},
                          AD_Agentur: contactIdADAgentur
                            ? { id: contactIdADAgentur }
                            : {},
                          Sanierungspartner: contactIdSanierungspartner
                            ? { id: contactIdSanierungspartner }
                            : contactIdSanierungspartner,
                          Schadenursache: fileData?.top?.["Ursache: "],
                        },
                        Trigger: ["workflow"],
                      };
                      await ZOHO.CRM.API.updateRecord(config);
                      console.log(config);

                      // add note
                      const newLine = String.fromCharCode(0x0a);

                      const addNoteResp = await ZOHO.CRM.API.addNotes({
                        Entity: "Deals",
                        RecordID: deal?.id,
                        Title: "Vertragsdaten Data",
                        Content:
                          "Vertragsnummer: " +
                          fileData?.bottom?.["Vertragsnummer:"] +
                          newLine +
                          "Versicherte Sache / Risiko: " +
                          fileData?.bottom?.["Versicherte Sache / Risiko:"] +
                          newLine +
                          "Versicherungsbedingungen: " +
                          fileData?.bottom?.["Versicherungsbedingungen:"] +
                          newLine +
                          "Reserve: " +
                          fileData?.bottom?.["Reserve:"] +
                          newLine +
                          "Versicherungssumme: " +
                          fileData?.bottom?.["Versicherungssumme:"] +
                          newLine +
                          "Euro: " +
                          fileData?.bottom?.["Euro:"] +
                          newLine +
                          "Baujahr: " +
                          fileData?.bottom?.["Baujahr:"] +
                          newLine +
                          "Wohnfläche: " +
                          fileData?.bottom?.["Wohnfläche:"] +
                          newLine +
                          "Staffelentschädigung: " +
                          fileData?.bottom?.["Staffelentschädigung:"] +
                          newLine +
                          "letzte Schätzung: " +
                          fileData?.bottom?.["letzte Schätzung:"] +
                          newLine +
                          "Selbstbehalt: " +
                          fileData?.bottom?.["Selbstbehalt:"] +
                          newLine +
                          "Weitere wichtige Informationen zum Vertrag / Schaden: " +
                          fileData?.bottom?.[
                            "Weitere wichtige Informationen zum Vertrag / Schaden:"
                          ] +
                          newLine +
                          "Termin_Info: " +
                          fileData?.bottom?.["Termin_Info"],
                      });

                      console.log(addNoteResp);

                      // collect promises
                      const uploadPromises = filesArray.map((file, index) => {
                        return ZOHO.CRM.API.attachFile({
                          Entity: "Deals",
                          RecordID: deal.id,
                          File: {
                            Name: file.name,
                            Content: file,
                          },
                        }).then((res) => {
                          console.log("Uploaded:", res);
                          return res; // keep the result
                        });
                      });

                      // wait for all to finish
                      const results = await Promise.all(uploadPromises);

                      console.log("All uploads done", results);

                      // take decision after everything is uploaded
                      if (results.every((r) => r.data[0].code === "SUCCESS")) {
                        ///////////////////// upload all the attachments to workdrive and delete from zoho crm attachments section /////////////
                        //  deal_id
                        var func_name = "upload_to_workdrive_delete_from_rec";
                        var req_data = {
                          arguments: JSON.stringify({
                            deal_id: deal.id,
                          }),
                        };
                        await ZOHO.CRM.FUNCTIONS.execute(func_name, req_data);
                        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        console.log("All files uploaded successfully!");
                        const updatedFileNames = Object.keys(
                          allFilesData
                        ).filter((name) => name !== fileName);
                        setSelectedDeals((prev) => {
                          const updatedData = {};
                          updatedFileNames.forEach((name) => {
                            updatedData[name] = selectedDeals[name];
                          });
                          return updatedData;
                        });
                        setSelectedFiles((prev) => {
                          const updatedData = {};
                          updatedFileNames.forEach((name) => {
                            updatedData[name] = selectedFiles[name];
                          });
                          return updatedData;
                        });

                        setAllFilesData((prev) => {
                          const updatedData = {};
                          updatedFileNames.forEach((name) => {
                            updatedData[name] = allFilesData[name];
                          });
                          return updatedData;
                        });

                        // your next decision here
                      } else {
                        console.log("Some uploads failed", results);
                      }
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "10px",
                        flex: 1,
                      }}
                    >
                      <div
                        style={{
                          width: "450px",
                        }}
                      >
                        <Autocomplete
                          disablePortal
                          id="combo-box-demo"
                          onOpen={() => {
                            const data = {};
                            deals.forEach((option) => {
                              data[option.id] = option;
                            });

                            setDeals(Object.values(data));
                          }}
                          options={deals}
                          getOptionLabel={(option) => option?.Deal_Name}
                          getOptionKey={(option) => option.id}
                          sx={{
                            width: "100%",
                          }}
                          size={"small"}
                          value={
                            selectedDeals[fileName]
                              ? {
                                  Deal_Name:
                                    selectedDeals?.[fileName]?.name || "",
                                  id: selectedDeals?.[fileName]?.id || "",
                                }
                              : null
                          }
                          onChange={(event, value) => {
                            console.log({ value, id: value?.id });
                            setSelectedDeals((prev) => ({
                              ...prev,
                              [fileName]: {
                                name: value?.Deal_Name,
                                id: value?.id,
                              },
                            }));
                            // setValue(
                            //   {
                            //     name: value?.Full_Name,
                            //     id: value?.id,
                            //   } || ""
                            // );
                          }}
                          loading={loading}
                          loadingText={"Loading..."}
                          renderInput={(params) => (
                            <TextField
                              {...params}
                              required
                              sx={{ width: "100%" }}
                              placeholder="Select Deal Name"
                              onChange={async (e) => {
                                if (e?.target?.value?.length >= 3) {
                                  setSearchValue((prev) => e?.target?.value);
                                }
                              }}
                              // error={!!error}
                              // helperText={error && error.message}
                            />
                          )}
                        />
                      </div>

                      <input
                        type="file"
                        multiple
                        onChange={(e) => {
                          console.log(e.target.files);
                          setSelectedFiles((prev) => ({
                            ...prev,
                            [fileName]: e.target.files,
                          }));
                          // const name = e.target.files[0].name;

                          // ZOHO.CRM.API.attachFile({
                          //   Entity: "Deals",
                          //   RecordID: "4731441000014144866",
                          //   File: {
                          //     Name: name,
                          //     Content: e.target.files[0],
                          //   },
                          // }).then(function (data) {
                          //   console.log(data);
                          // });
                        }}
                      />

                      <Button type="submit" variant="contained" size="small">
                        Proceed
                      </Button>
                    </div>
                  </form>
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
      </Box>
    </Box>
  );
}
