const express = require("express");
const morgan = require("morgan");
const cors = require("cors");
const exphbs = require("express-handlebars");
const http = require("http");
const path = require("path");
const bodyParser = require("body-parser");
const app = express();
const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");
var nodemailer = require("nodemailer");
require("dotenv").config();

const server = http.createServer(app);
// const io = require("socket.io")(server);

const JWT_SECRET =
  'sdjkfh8923yhjdksbfmad3939&"#?"?#(#>Q(()@_#(##hjb2qiuhesdbhjdsfg839ujkdhfjk';

// require("./database");
app.set("views", path.join(__dirname, "views"));

const hbs = exphbs.create({
  defaultLayout: "main",
  layoutsDir: path.join(app.get("views"), "layouts"),
  partialsDir: path.join(app.get("views"), "partials"),
  extname: ".hbs",
  helpers: {
    ifeq: function (a, b, options) {
      if (a == b) {
        return options.fn(this);
      }
      return options.inverse(this);
    },
    ifnoteq: function (a, b, options) {
      if (a != b) {
        return options.fn(this);
      }
      return options.inverse(this);
    },
    firstL: function (options) {
      return options.charAt(0);
    },
  },
});

app.engine(".hbs", hbs.engine);
app.set("view engine", ".hbs");

// Middleware
app.use(bodyParser.json());
app.use(morgan("tiny")); //Morgan
app.use(cors()); // cors
app.use(express.json()); // JSON
app.use(express.urlencoded({ extended: false })); //urlencoded
app.use(bodyParser.json());

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

function send_email(resp, res, type, amount, mail) {
  let mailOption_client = {
    from: `GGF Internal <ggf_internal@myguardiangroup.com>`,
    to: mail,
    subject: "Qr-Code code",
    body: resp,
    attachments: [
      {
        filename: "qr_codes.xlsx",
        path: resp,
        contentType: "application/xlsx",
      },
    ],
    html: `<!DOCTYPE html>
  <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" lang="en">
  
  <head>
    <title></title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      * {
        box-sizing: border-box;
      }
  
      body {
        margin: 0;
        padding: 0;
      }
  
      a[x-apple-data-detectors] {
        color: inherit !important;
        text-decoration: inherit !important;
      }
  
      #MessageViewBody a {
        color: inherit;
        text-decoration: none;
      }
  
      p {
        line-height: inherit
      }
  
      .desktop_hide,
      .desktop_hide table {
        mso-hide: all;
        display: none;
        max-height: 0px;
        overflow: hidden;
      }
  
      .image_block img+div {
        display: none;
      }
  
      @media (max-width:655px) {
  
        .desktop_hide table.icons-inner,
        .social_block.desktop_hide .social-table {
          display: inline-block !important;
        }
  
        .icons-inner {
          text-align: center;
        }
  
        .icons-inner td {
          margin: 0 auto;
        }
  
        .mobile_hide {
          display: none;
        }
  
        .row-content {
          width: 100% !important;
        }
  
        .stack .column {
          width: 100%;
          display: block;
        }
  
        .mobile_hide {
          min-height: 0;
          max-height: 0;
          max-width: 0;
          overflow: hidden;
          font-size: 0px;
        }
  
        .desktop_hide,
        .desktop_hide table {
          display: table !important;
          max-height: none !important;
        }
      }
    </style>
  </head>
  
  <body class="body" style="margin: 0; background-color: #f1f1f1; padding: 0; -webkit-text-size-adjust: none; text-size-adjust: none;">
    <table class="nl-container" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f1f1f1;">
      <tbody>
        <tr>
          <td>
            <table class="row row-1" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
              <tbody>
                <tr>
                  <td>
                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f1f1f1; color: #000000; width: 635px; margin: 0 auto;" width="635">
                      <tbody>
                        <tr>
                          <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-bottom: 5px; padding-top: 5px; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                            <div class="spacer_block block-1" style="height:20px;line-height:20px;font-size:1px;">&#8202;</div>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table class="row row-2" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
              <tbody>
                <tr>
                  <td>
                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #ffffff; color: #000000; width: 635px; margin: 0 auto;" width="635">
                      <tbody>
                        <tr>
                          <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-bottom: 15px; padding-top: 15px; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                            <table class="image_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                              <tr>
                                <td class="pad" style="width:100%;padding-right:0px;padding-left:0px;">
                                  <div class="alignment" align="center" style="line-height:10px">
                                    <div style="max-width: 95.25px;"><img src="https://d9638f13c3.imgdist.com/pub/bfra/58t8elma/mo8/5ey/2e6/guardian_logo_1.png" style="display: block; height: auto; border: 0; width: 100%;" width="95.25" alt="Logo" title="Logo" height="auto"></div>
                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table class="row row-3" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
              <tbody>
                <tr>
                  <td>
                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #ffffff; color: #000000; width: 635px; margin: 0 650px;" width="635">
                      <tbody>
                        <tr>
                          <td class="column column-1" style="width:50%;mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-bottom: 5px;padding-bottom: 5px;padding-left: 28px; padding-top: 10px; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                            <table class="paragraph_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-top:10px;">
                                  <div style="color:#555555;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:14px;line-height:120%;text-align:left;mso-line-height-alt:16.8px;">&nbsp;</div>
                                </td>
                              </tr>
                            </table>
                            <table class="paragraph_block block-2" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-bottom:10px;padding-top:10px;">
                                  <div style="color:#9c0059;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:30px;line-height:120%;text-align:left;mso-line-height-alt:36px;">
                                    <p style="margin: 0; word-break: break-word;"><span><strong>Hi, <br>In attachment you will find and excel file with all the qr-code codes</strong></span></p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="paragraph_block block-3" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-top:5px;">
                                  <div style="color:#9c0059;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:13px;line-height:120%;text-align:left;mso-line-height-alt:15.6px;">
                                    <p style="margin: 0; word-break: break-word;"><span>${type} qr-code</span></p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="text_block block-4" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-bottom:10px;padding-top:5px;">
                                  <div style="font-family: sans-serif">
                                    <div class style="font-size: 14px; font-family: Poppins, Arial, Helvetica, sans-serif; mso-line-height-alt: 28px; color: #808080; line-height: 2;">
                                      <p style="margin: 0; font-size: 14px; mso-line-height-alt: 32px;"><span style="font-size:14px;">${amount} qr-code in result</span></p>
                                      <p style="margin: 0; font-size: 14px; mso-line-height-alt: 28px;">&nbsp;</p>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                          <td class="column column-2"  style="width:40%;mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-top: 10px; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                            <table class="image_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                              <tr>
                                <td class="pad" style="width:100%;">
                                  <div class="alignment" align="right" style="line-height:10px">
                                    <div style="max-width: 254px;"><img src="https://d1oco4z2z1fhwp.cloudfront.net/templates/default/1126/featured-image.png" style="display: block; height: auto; border: 0; width: 100%;" width="254" alt="Image" title="Image" height="auto"></div>
                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table class="row row-4" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
              <tbody>
                <tr>
                  <td>
                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #f1f1f1; color: #000000; width: 635px; margin: 0 auto;" width="635">
                      <tbody>
                        <tr>
                          <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-bottom: 15px; padding-top: 15px; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                            <div class="spacer_block block-1" style="height:20px;line-height:20px;font-size:1px;">&#8202;</div>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
            <table class="row row-5" align="center" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #322985;">
              <tbody>
                <tr>
                  <td>
                    <table class="row-content stack" align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; background-color: #322985; color: #000000; width: 635px; margin: 0 auto;" width="635">
                      <tbody>
                        <tr>
                          <td class="column column-1" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; font-weight: 400; text-align: left; padding-bottom: 20px; padding-top: 10px; vertical-align: top; border-top: 0px; border-right: 0px; border-bottom: 0px; border-left: 0px;">
                            <table class="paragraph_block block-1" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-bottom:5px;padding-left:10px;padding-right:10px;padding-top:25px;">
                                  <div style="color:#ffffff;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:18px;line-height:120%;text-align:center;mso-line-height-alt:21.599999999999998px;">
                                    <p style="margin: 0; word-break: break-word;"><span>FOLLOW US:</span></p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="social_block block-2" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                              <tr>
                                <td class="pad">
                                  <div class="alignment" align="center">
                                    <table class="social-table" width="138px" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; display: inline-block;">
                                      <tr>
                                        <td style="padding:0 7px 0 7px;"><a href="https://www.facebook.com/GuardianGroupDutchCaribbean" target="_blank"><img src="https://app-rsrc.getbee.io/public/resources/social-networks-icon-sets/t-circle-white/facebook@2x.png" width="32" height="auto" alt="Facebook" title="Facebook" style="display: block; height: auto; border: 0;"></a></td>
                                        <td style="padding:0 7px 0 7px;"><a href="https://www.instagram.com/guardiangroupdc/" target="_blank"><img src="https://app-rsrc.getbee.io/public/resources/social-networks-icon-sets/t-circle-white/instagram@2x.png" width="32" height="auto" alt="Instagram" title="Instagram" style="display: block; height: auto; border: 0;"></a></td>
                                        <td style="padding:0 7px 0 7px;"><a href="https://www.linkedin.com/company/my-guardian-group/mycompany/" target="_blank"><img src="https://app-rsrc.getbee.io/public/resources/social-networks-icon-sets/t-circle-white/linkedin@2x.png" width="32" height="auto" alt="LinkedIn" title="LinkedIn" style="display: block; height: auto; border: 0;"></a></td>
                                      </tr>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="paragraph_block block-3" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad">
                                  <div style="color:#ffffff;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:14px;line-height:180%;text-align:center;mso-line-height-alt:25.2px;">
                                    <p style="margin: 0; word-break: break-word;"><a style="text-decoration: none; color: #ffffff;" title="tel:+12025550109" href="tel:+12025550109">+599 9 777 7777</a></p>
                                    <p style="margin: 0;">www.portal.myguardiangroup.com/en</p>
                                    <p style="margin: 0;">2 Cas Coraweg, Willemstad, Curacao</p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="divider_block block-4" width="100%" border="0" cellpadding="10" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                              <tr>
                                <td class="pad">
                                  <div class="alignment" align="center">
                                    <table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt;">
                                      <tr>
                                        <td class="divider_inner" style="font-size: 1px; line-height: 1px; border-top: 1px solid #93CADE;"><span>&#8202;</span></td>
                                      </tr>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="paragraph_block block-5" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-left:10px;padding-right:10px;">
                                  <div style="color:#ffffff;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:14px;line-height:120%;text-align:center;mso-line-height-alt:16.8px;">
                                    <p style="margin: 0; word-break: break-word;">&nbsp;</p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="paragraph_block block-6" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-left:10px;padding-right:10px;">
                                  <div style="color:#ffffff;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:11px;line-height:120%;text-align:center;mso-line-height-alt:13.2px;">
                                    <p style="margin: 0; word-break: break-word;"><span>Thank you for using our Qr-code generator. If you have any question please send us an email at</span></p>
                                    <p style="margin: 0; word-break: break-word;"><span>customerservice@myguardiangroup.com and we will help you with anything you need</span></p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                            <table class="paragraph_block block-7" width="100%" border="0" cellpadding="0" cellspacing="0" role="presentation" style="mso-table-lspace: 0pt; mso-table-rspace: 0pt; word-break: break-word;">
                              <tr>
                                <td class="pad" style="padding-left:10px;padding-right:10px;">
                                  <div style="color:#ffffff;font-family:Poppins, Arial, Helvetica, sans-serif;font-size:14px;line-height:120%;text-align:center;mso-line-height-alt:16.8px;">
                                    <p style="margin: 0; word-break: break-word;">&nbsp;</p>
                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </td>
                </tr>
              </tbody>
            </table>
          </td>
        </tr>
      </tbody>
    </table><!-- End -->
  </body>
  
  </html>`,
  };

  //send email to client
  let transporter_client = nodemailer.createTransport({
    host: "smtp.sendgrid.net",
    port: 587,
    secure: false,
    auth: {
      user: "apikey",
      pass: process.env.EMAIL_P,
    },
    tls: {
      rejectUnauthorized: false,
    },
  });

  transporter_client.sendMail(mailOption_client, function (error, info) {
    if (error) {
      console.error("Error sending email:", error);
      res.status(500).json({ status: "404", error: error.message });
    } else {
      console.log("Email sent: " + info.response);
      res.status(202).json({ msg: "Successful!", status: "202" });
    }
  });
}

app.use(require("./routes"));
app.use(express.static(path.join(__dirname, "public")));

// const server = http.createServer(app);

app.set("port", process.env.PORT || 8085);

server.listen(app.get("port"), () => {
  console.log("server on port", app.get("port"));
});
