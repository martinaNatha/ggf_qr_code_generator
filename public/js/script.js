var jsonData;

$("#singl_sub").on("click", function () {
  var idnum = document.getElementById("idnum").value;
  // var mail = document.getElementById("mail").value;
  send_data(idnum, "Single");
});

$("#idnum").on("change", function () {
  const value = $(this).val();
  console.log(value);
  if (value.length == 10) {
    setTimeout(() => {
      document.getElementById("singl_sub").style.display = "block";
      document.getElementById("singl_sub").classList.add("fade-in");
    }, 500);
  } else {
  }
});

document.getElementById("excelwninput").addEventListener("change", multiqr);

function multiqr(event) {
  const file = event.target.files[0];

  // Check if the file type is not Excel
  const validExtensions = ["xls", "xlsx"];
  const fileExtension = file.name.split(".").pop().toLowerCase();
  if (!validExtensions.includes(fileExtension)) {
    Swal.fire({
      title: "Invalid File Type",
      text: "Please upload a valid Excel file (.xls or .xlsx).",
      icon: "error",
    });
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Define the column letter or index from which you want to extract data
    const targetColumn = "4"; // For example, extracting data from column D

    // Find the range of the data
    const range = XLSX.utils.decode_range(worksheet["!ref"]);

    // Check if headers are at the specific position
    const expectedHeaders = [
      "VERZEKERDE",
      "VERZEKERINGNEMER",
      "MEDIFLEXPOLIS",
      "POLISNUMMER",
      "IDNUMMER",
      "GELDIG T/M",
      "GESLACHT",
    ];
    const actualHeaders = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: c });
      const cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : "";
      actualHeaders.push(cellValue);
    }

    const headersMatch = expectedHeaders.every(
      (header, index) => header === actualHeaders[index]
    );

    if (!headersMatch) {
      console.error("Headers do not match the expected format.");
      Swal.fire({
        title: "Error",
        text: "The headers of the uploaded Excel file do not match the expected format. Please upload a file with the correct headers.",
        icon: "error",
      });
      return;
    } else {
      document.getElementById("Emessage").style.display = "block";
      document.getElementById("Emessage").classList.add("fade-in");
      setTimeout(() => {
        document.getElementById("multi_sub").style.display = "block";
        document.getElementById("multi_sub").classList.add("fade-in");
      }, 500);
    }

    // Start reading from the second row (excluding the header)
    range.s.r++;

    // Extract data from the target column (excluding the header row)
    const columnData = [];
    for (let i = range.s.r; i <= range.e.r; i++) {
      const cellAddress = XLSX.utils.encode_cell({ r: i, c: targetColumn - 1 });
      const cellValue = worksheet[cellAddress] ? worksheet[cellAddress].v : ""; // Get cell value (or empty string if cell is empty)
      columnData.push(cellValue);
    }

    // Add the extracted data from the column to a list
    const dataList = columnData.filter(Boolean); // Filter out empty values

    // Display the list in the console
    jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false });
    // send_data(dataList);
    console.log(jsonData);
  };

  reader.readAsArrayBuffer(file);
}

$("#multi_sub").on("click", function () {
  // send_data(jsonData, "Multiple");
  var preload = document.getElementById("loader");
  var response_msg2 = document.getElementById("email2");
  setTimeout(function () {
    preload.style.display = "flex";
    preload.style.opacity = 1;
    response_msg2.style.display = "flex";
    response_msg2.style.opacity = 1;
  }, 1000);
});

$("#mail_sub").on("click", function () {
  var mail = document.getElementById("mail").value;
  send_data(jsonData, "Multiple", mail);
});

async function send_data(data, type, mail) {
  var loader_15 = document.getElementById("loader-15");
  var resp_msg = document.getElementById("resp_msg");
  var email2 = document.getElementById("email2");
  email2.style.display = "none";
  email2.style.opacity = 0;
  setTimeout(function () {
    loader_15.style.display = "flex";
    loader_15.style.opacity = 1;
  }, 1000);
  var data_body;
  if (type == "Single") {
    data_body = `[{"sedula":"` + data + `"}]`;
  } else {
    var data_body = data.map((row) => {
      return { sedula: row.IDNUMMER };
    });
    console.log(data_body);
  }
  const result = await fetch("/get_qr_code", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      data: data_body,
      data2: data,
      type: type,
      email: mail,
    }),
  }).then((res) => res.json());
  if (result.status == "202") {
    if (type == "Single") {
      var single_response = document.getElementById("single_response");
      setTimeout(function () {
        single_response.style.display = "flex";
        single_response.style.opacity = 1;
      }, 1000);
      document.getElementById("single_response").innerText =
        "ID NUMBER CODE: " + result.resp;
    } else {
      setTimeout(function () {
        loader_15.style.display = "none";
        loader_15.style.opacity = 0;
        resp_msg.style.display = "flex";
        resp_msg.style.opacity = 1;
      }, 1000);
    }
  }
}

$("#close_b").on("click", function () {
  var preload = document.getElementById("loader");
  var resp_msg = document.getElementById("resp_msg");
  preload.style.display = "none";
  preload.style.opacity = 0;
  resp_msg.style.display = "none";
  resp_msg.style.opacity = 0;
});

const alertButton = document.getElementById("info_b");
  const customAlert = document.getElementById("customAlert");
  const closeBtn = document.querySelector(".close-btn");

  alertButton.addEventListener("mouseenter", function() {
    customAlert.style.display = "flex"; // Show the alert box
    setTimeout(() => {
      customAlert.classList.add("show"); // Trigger the slide-down effect
    }, 10); // Slight delay to trigger the transition
  });

  // Optionally, hide the alert box when clicking outside of it
  window.addEventListener("click", function(event) {
    if (event.target === customAlert) {
      customAlert.classList.remove("show");
      setTimeout(() => {
        customAlert.style.display = "none";
      }, 300);
    }
  });