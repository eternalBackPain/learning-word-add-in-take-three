/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { base64Image } from "../../base64Image";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("replace-text").onclick = replaceText;
    document.getElementById("insert-image").onclick = insertImage;
    document.getElementById("insert-html").onclick = insertHTML;
    document.getElementById("insert-table").onclick = insertTable;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function insertParagraph() {
  await Word.run(async (context) => {
    // Queue commands to insert a paragraph into the document.
    const docBody = context.document.body;
    //insertParagraph method takes two arguments: the text to insert, and the location where the paragraph should be inserted
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
      "Start"
    );
    await context.sync(); // sends all queued commands to Word for execution
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    // Queue commands to style text.
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {
    // Queue commands to apply the custom style.
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function changeFont() {
  await Word.run(async (context) => {
    // Queue commands to apply a different font. Code gets a reference to the second paragraph by using the ParagraphCollection.getFirst method chained to the Paragraph.getNext method.
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {
    // Queue commands to insert text into a selected range.
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");

    //Add code to fetch document properties into the task pane's script objects:
    // Load the text of the range and sync so that the current range text can be read.
    originalRange.load("text");
    await context.sync();

    // Queue commands to repeat the text of the original range at the end of the document. We can now read .text because of our above load and sync.
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");

    // Move the final call of context.sync here and ensure that it does not run until the insertParagraph has been queued.
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {
    // Queue commands to insert a new range before the selected range.
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");

    // Load the text of the original range and sync so that the range text can be read and inserted.
    originalRange.load("text");
    await context.sync();

    // Queue commands to insert the original range as a paragraph at the end of the document - THE ORIGINAL RANGE DOESNT CHANGE WHEN INSERTING TEXT
    doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");

    // Make a final call of context.sync here and ensure that it runs after the insertParagraph has been queued.
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function replaceText() {
  await Word.run(async (context) => {
    // Queue commands to replace the text.
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertImage() {
  await Word.run(async (context) => {
    // Queue commands to insert an image.
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    //Note: The Paragraph object also has an insertInlinePictureFromBase64 method and other insert* methods. See the following insertHTML section for an example.
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertHTML() {
  await Word.run(async (context) => {
    // adds a blank paragraph to the end of the document
    const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");

    //inserts a string of HTML at the end of the paragraph
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function insertTable() {
  await Word.run(async (context) => {
    // get a reference to the paragraph that will proceed the table.
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

    // create a table and populate it with data.
    const tableData = [
      ["Name", "ID", "Birth City"],
      ["Bob", "434", "Chicago"],
      ["Sue", "719", "Havana"],
    ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
