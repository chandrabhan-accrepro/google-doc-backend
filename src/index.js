const express = require("express");
const app = express();
const cors = require("cors");

const { google } = require("googleapis");

const multer = require("multer");
const fs = require("fs");
const path = require("path");
const axios = require("axios");
const port = process.env.PORT || 5000;

const upload = multer({ storage: multer.memoryStorage() });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
const allowedOrigins = [
  "http://localhost:3000",
  "https://google-doc-integration-front.onrender.com/",
];
app.use(
  cors({
    origin: (origin, callback) => {
      // Check if the origin is in the allowedOrigins array or if it is undefined (for local development)
      if (allowedOrigins.indexOf(origin) !== -1 || !origin) {
        callback(null, true);
      } else {
        callback(new Error("Not allowed by CORS"));
      }
    },
  })
);

app.get("/create", async (req, res) => {
  try {
    const SCOPES = ["https://www.googleapis.com/auth/drive"];
    const credentials = require("./account-key.json"); // Replace with your own credentials file

    const client = new google.auth.JWT(
      credentials.client_email,
      null,
      credentials.private_key,
      SCOPES
    );
    const drive = google.drive({ version: "v3", auth: client });

    const fileMetadata = {
      name: "My New Document",
      mimeType: "application/vnd.google-apps.document",
    };

    const response = await drive.files.create({
      resource: fileMetadata,
    });
    const permission = await drive.permissions.create({
      fileId: response.data.id,
      resource: {
        role: "writer",
        type: "anyone",
      },
    });
    // console.log("Permission added: ", permission);
    console.log("Creation FileID: ", response.data);
    const documentUrl = `https://docs.google.com/document/d/${response.data.id}/edit?usp=drivesdk`;
    // const documentUrl = `https://docs.google.com/document/d/1aFMMugFsjJ0oUp1ZC_JtDocuq0wDJkFSNWF67I7bZ4o/edit?usp=drivesdk`;
    res.status(200).json({
      message: "Successfully Created a Document",
      documentId: response.data.id,
      documentUrl,
      fileName: response.data.name,
    });
  } catch (err) {
    res.status(500).send({ message: err.message });
  }
});

// Convert and edit the uploaded document
app.post("/editDocument", upload.single("file"), async (req, res) => {
  console.log("Authent");
  try {
    const uploadedFile = req?.file?.buffer;
    const pathOfFile = {};
    if (uploadedFile) {
      console.log(req.file);
      const originalFileName = req.file.originalname;

      // Define the path where you want to save the uploaded file
      const savePath = path.join(__dirname, "docs", originalFileName);
      pathOfFile.savePath = savePath;
      pathOfFile.originalFileName = originalFileName;
      console.log(savePath);
      // Save the uploaded file to your local storage
      await fs.writeFile(savePath, uploadedFile, (err) => {
        if (err) {
          console.error("Error saving uploaded file:", err);
        } else {
          console.log("File saved successfully:", savePath);
        }
      });
    } else {
      const originalFileName = req.query.fileName;
      console.log("FileName:", req.query.fileName);
      // Define the path where you want to save the uploaded file
      pathOfFile.savePath = path.join(__dirname, "editted", originalFileName);
      pathOfFile.originalFileName = originalFileName;
    }
    // Create a new Google Docs document
    const savePath = pathOfFile.savePath;
    const originalFileName = pathOfFile.originalFileName;
    console.log("Checking the path:", savePath);
    const SCOPES = ["https://www.googleapis.com/auth/drive"];
    const credentials = require("./account-key.json"); // Replace with your own credentials file

    const client = new google.auth.JWT(
      credentials.client_email,
      null,
      credentials.private_key,
      SCOPES
    );
    const drive = google.drive({ version: "v3", auth: client });
    const fileMetadata = {
      name: originalFileName,
      mimeType: "application/vnd.google-apps.document", // Google Docs format
    };

    const media = {
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // Word document MIME type
      body: fs.createReadStream(savePath),
    };
    let fileurl = {};

    try {
      const file = await drive.files.create({
        resource: fileMetadata,
        media: media,
        fields: "id,name,mimeType,webViewLink,parents",
      });
      fileurl.url = file.data.webViewLink;
      fileurl.fileid = file.data.id;
      fileurl.name = file.data.name;
      console.log("Google Docs File ID:", file.data);
      const permission = await drive.permissions.create({
        fileId: file.data.id,
        resource: {
          role: "writer",
          type: "anyone",
        },
      });
      console.log("Permission added: ", permission);
      const getFile = await drive.files.export({
        fileId: file.data.id,
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      console.log(getFile);
      // return file.data.id;
    } catch (err) {
      // TODO (developer) - Handle error
      throw err;
    }

    const documentUrl = fileurl.url;
    // const documentUrl = `https://docs.google.com/document/d/1kLO96MgOx7eU7E9BLnjmMvjkf6v-yBaFluUuRNlL110/edit`;

    res.status(200).json({
      documentUrl,
      fileId: fileurl.fileid,
      fileName: fileurl.name,
    });
  } catch (err) {
    console.log(err);
    res.status(401).json({ message: err });
  }
});

app.get("/savedoc", async (req, res) => {
  console.log("saved data");

  try {
    // const auth = oauth2Client;
    console.log("documentID", req.query.fileId);

    console.log("Path of key: ", path.join(__dirname, "account-key.json"));

    try {
      const SCOPES = ["https://www.googleapis.com/auth/drive"];
      const credentials = require("./account-key.json"); // Replace with your own credentials file

      const client = new google.auth.JWT(
        credentials.client_email,
        null,
        credentials.private_key,
        SCOPES
      );
      const drive = await google.drive({ version: "v3", auth: client });

      const exportMimeType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; // Change this to the desired export format

      const findFile = await drive.files.get({ fileId: req.query.fileId });
      console.log(findFile.data.name);

      drive.files.export(
        {
          fileId: req.query.fileId,
          mimeType: exportMimeType,
        },
        { responseType: "stream" },
        (exportErr, responses) => {
          if (exportErr) {
            console.error("Error exporting file:", exportErr);
            return;
          }
          console.log("Saving the document: ", responses);
          console.log("FileName checking: ", req.query.fileName);
          // Handle the exported file stream (e.g., save it to your server)
          const outputPath = path.join(
            __dirname,
            "editted",
            `${findFile.data.name}_${req.query.fileId}.docx`
          );
          responses.data.pipe(fs.createWriteStream(outputPath));
          res.status(200).json({ message: "Successfully Saved Document" });
        }
      );
    } catch (error) {
      console.error("Error exporting Google Doc as PDF:", error);
      res.status(500).json({ message: "Error exporting Google Doc as PDF" });
    }

    // res.status(200).send("Word document saved successfully.");
    // console.log(util.inspect(response.data, false, 17));
    // console.log("Response data: ", JSON.stringify(response.data));
  } catch (err) {
    console.log(err);
  }
});

app.get("/api/v1/getPdfs", (req, res) => {
  try {
    const pdfFolder = path.join(__dirname, "editted"); // Replace with your folder path

    // Read all files in the PDF folder
    fs.readdir(pdfFolder, (err, files) => {
      if (err) {
        // Handle any error that occurs while reading the folder
        console.error(err.message);
        res.status(500).send("Internal Server Error");
      } else {
        // Filter only PDF files
        const pdfFiles = files.filter((file) =>
          file.toLowerCase().endsWith(".docx")
        );

        // Set the Content-Type header to indicate that it's a PDF file
        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );
        // console.log(pdfFiles);
        // Send all PDF files as a response
        const allLinks = [];
        pdfFiles.forEach((filename) => {
          const filePath = path.join(pdfFolder, filename);
          // console.log(filePath);
          allLinks.push(
            filename
            // `<a href="http://localhost:8000/api/v1/getPdf/${filename}">${filename}</a><br>`
          );
        });
        console.log(allLinks);
        res.status(200).json({ result: allLinks });
      }
    });
  } catch (error) {
    console.log("error: ", error.message);
    res.status(500).json({ error: error });
  }
});

app.get("/api/v1/getPdf/:filename", (req, res) => {
  const pdfFolder = path.join(__dirname, "editted");
  const { filename } = req.params;
  const filePath = path.join(pdfFolder, filename);
  console.log(pdfFolder);
  // Check if the file exists
  if (fs.existsSync(filePath)) {
    // Set the Content-Type header to indicate that it's a PDF file
    res.setHeader("Content-Type", "application/pdf");

    // Send the PDF file as a response
    console.log(filePath);
    res.sendFile(filePath);
  } else {
    // Handle the case where the file does not exist
    res.status(404).send("File not found");
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
