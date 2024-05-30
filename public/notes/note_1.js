app.post("/get_qr_code", async (req, res) => {
    var { data, data2, type, email } = req.body;
  
    if (type == "Multiple") {
      const outputDir = path.resolve(__dirname, "excel_sheets");
  
      // Function to ensure the directory exists
      function ensureDirectoryExistence(directory) {
        if (!fs.existsSync(directory)) {
          fs.mkdirSync(directory, { recursive: true });
          console.log(`Directory created: ${directory}`);
        } else {
          console.log(`Directory already exists: ${directory}`);
        }
      }
  
      // Ensure the directory exists before proceeding
      ensureDirectoryExistence(outputDir);
  
      // Check if data is an array and has more than one item
      if (Array.isArray(data) && data.length > 1) {
        const jsonData = JSON.stringify(data);
        const jsonData2 = data2;
  
        axios
          .post("https://misalu.live/ords/api/ggf-pasjes/nonce", jsonData, {
            headers: {
              "Content-Type": "application/json",
              "API-KEY": "GGFTenBehoeveVanPasjes2022",
            },
          })
          .then((response) => {
            // Check if response.data is a string and split it by newline characters
            const entries =
              typeof response.data === "string"
                ? response.data.split("\n")
                : response.data;
  
            // Assuming response.data is an array of new values
            const newValues = entries;
  
            // Convert JSON to worksheet (assuming jsonData2 is already defined)
            const worksheet = XLSX.utils.json_to_sheet(jsonData2);
  
            // Get the range of the worksheet
            const range = XLSX.utils.decode_range(worksheet["!ref"]);
  
            // Add a new header for the additional column
            worksheet[XLSX.utils.encode_cell({ r: 0, c: range.e.c + 1 })] = {
              v: "CODE",
            };
  
            // Add new values to the additional column
            for (let i = 0; i < newValues.length; i++) {
              worksheet[XLSX.utils.encode_cell({ r: i + 1, c: range.e.c + 1 })] =
                {
                  v: newValues[i],
                };
            }
  
            // Update the range to include the new column
            range.e.c++;
            worksheet["!ref"] = XLSX.utils.encode_range(range);
  
            // Create a new workbook and add the worksheet to it
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  
            // Write the workbook to a file
            const now = new Date();
  
            const year = now.getFullYear();
            const month = String(now.getMonth() + 1).padStart(2, "0"); // Months are 0-based
            const day = String(now.getDate()).padStart(2, "0");
            const hours = String(now.getHours()).padStart(2, "0");
            const minutes = String(now.getMinutes()).padStart(2, "0");
            const seconds = String(now.getSeconds()).padStart(2, "0");
  
            const fullDateTime = `${year}${month}${day}${hours}${minutes}${seconds}`;
            const outputFilePath = path.join(
              outputDir,
              `qrvb_` + fullDateTime + `.xlsx`
            );
            XLSX.writeFile(workbook, outputFilePath);
            send_email(outputFilePath, res, "Multiple", data.length, email);
            // console.log("Excel file generated successfully at", outputFilePath);
            res.status(202).json({ msg: "Successful!", status: "202" });
            return;
          })
          .catch((error) => {
            console.error("There was a problem with the request:", error);
          });
      } else {
        console.error("Data must be an array with more than one item.");
      }
    } else {
      console.log(data);
      axios
        .post("https://misalu.live/ords/api/ggf-pasjes/nonce", data, {
          headers: {
            "Content-Type": "application/json",
            "API-KEY": "GGFTenBehoeveVanPasjes2022",
          },
        })
        .then((response) => {
          console.log("API response:", response.data);
          res.json({ status: "202", resp: response.data });
          // send_email(response.data, res, "single", data.length, email);
        })
        .catch((error) => {
          console.error("There was a problem with the request:", error);
        });
    }
  });



  app.post("/get_qr_code", async (req, res) => {
    var { data, data2, type, email } = req.body;
  
    if (type == "Multiple") {
      const outputDir = path.resolve(__dirname, "excel_sheets");
    
      function ensureDirectoryExistence(directory) {
        if (!fs.existsSync(directory)) {
          fs.mkdirSync(directory, { recursive: true });
          console.log(`Directory created: ${directory}`);
        } else {
          console.log(`Directory already exists: ${directory}`);
        }
      }
  
      ensureDirectoryExistence(outputDir);
  
      const jsonData = JSON.stringify(data);
      const jsonData2 = data2;
  
      // Function to make API calls
      const makeApiCall = async (dataToSend) => {
        const response = await axios.post(
          "https://misalu.live/ords/api/ggf-pasjes/nonce",
          JSON.stringify(dataToSend),
          {
            headers: {
              "Content-Type": "application/json",
              "API-KEY": "GGFTenBehoeveVanPasjes2022",
            },
          }
        );
        return response.data;
      };
  
      const handleData = async () => {
        let entries = [];
  
        try {
          if (data.length > 10) {
            const midIndex = Math.ceil(data.length / 2);
            const firstHalf = data.slice(0, midIndex);
            const secondHalf = data.slice(midIndex);
  
            const [firstHalfResponse, secondHalfResponse] = await Promise.all([
              makeApiCall(firstHalf),
              makeApiCall(secondHalf),
            ]);
  
            entries = [
              ...(typeof firstHalfResponse === "string"
                ? firstHalfResponse.split("\n")
                : firstHalfResponse),
              ...(typeof secondHalfResponse === "string"
                ? secondHalfResponse.split("\n")
                : secondHalfResponse),
            ];
          } else {
            const response = await makeApiCall(data);
            entries =
              typeof response === "string" ? response.split("\n") : response;
          }
  
          const newValues = entries;
  
          const worksheet = XLSX.utils.json_to_sheet(jsonData2);
  
          
          const range = XLSX.utils.decode_range(worksheet["!ref"]);
  
          
          worksheet[XLSX.utils.encode_cell({ r: 0, c: range.e.c + 1 })] = {
            v: "CODE",
          };
  
          for (let i = 0; i < newValues.length; i++) {
            worksheet[XLSX.utils.encode_cell({ r: i + 1, c: range.e.c + 1 })] = {
              v: newValues[i],
            };
          }
  
          range.e.c++;
          worksheet["!ref"] = XLSX.utils.encode_range(range);
  
          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  
          const now = new Date();
  
          const year = now.getFullYear();
          const month = String(now.getMonth() + 1).padStart(2, "0");
          const day = String(now.getDate()).padStart(2, "0");
          const hours = String(now.getHours()).padStart(2, "0");
          const minutes = String(now.getMinutes()).padStart(2, "0");
          const seconds = String(now.getSeconds()).padStart(2, "0");
  
          const fullDateTime = `${year}${month}${day}${hours}${minutes}${seconds}`;
          const outputFilePath = path.join(
            outputDir,
            `qrvb_${fullDateTime}.xlsx`
          );
          XLSX.writeFile(workbook, outputFilePath);
  
          send_email(outputFilePath, res, "Multiple", data.length, email);
          res.status(202).json({ msg: "Successful!", status: "202" });
        } catch (error) {
          console.error("There was a problem with the request:", error);
          res.status(500).json({ msg: "Error", status: "500" });
        }
      };
  
      handleData();
    } else {
      console.log(data);
      axios
        .post("https://misalu.live/ords/api/ggf-pasjes/nonce", data, {
          headers: {
            "Content-Type": "application/json",
            "API-KEY": "GGFTenBehoeveVanPasjes2022",
          },
        })
        .then((response) => {
          console.log("API response:", response.data);
          res.json({ status: "202", resp: response.data });
          // send_email(response.data, res, "single", data.length, email);
        })
        .catch((error) => {
          console.error("There was a problem with the request:", error);
        });
    }
  });