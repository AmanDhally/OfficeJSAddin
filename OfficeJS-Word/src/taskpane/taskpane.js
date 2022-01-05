/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("footer").onclick = footerText;
    document.getElementById("footerImage").onclick = footerImage;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}
export async function HeaserAndfooter() {
  return Word.run(async (context) => {


    var sections = context.document.sections;
    context.load(sections, 'body/style');
    return context.sync().then(function () {
      var section = sections.items[0];
      //
      var primaryHeader = section.getHeader("primary");
      var firstPageHeader = section.getHeader("firstPage");
      var evenPagesHeader = section.getHeader("evenPages");
      var primaryFooter = section.getFooter("primary");
      var firstPageFooter = section.getFooter("firstPage");
      var evenPagesFooter = section.getFooter("evenPages");
      //
      primaryHeader.insertText("primary Header", Word.InsertLocation.start);
      firstPageHeader.insertText("first Page Header", Word.InsertLocation.start);
      evenPagesHeader.insertText("even Pages Header", Word.InsertLocation.start);
      primaryFooter.insertText("primary Footer", Word.InsertLocation.start);
      firstPageFooter.insertText("first Page Footer", Word.InsertLocation.start);
      evenPagesFooter.insertText("evenPages Footer", Word.InsertLocation.start);
      return context.sync().then(function () {
        // console.log("Completed");
      });
    });
  }).catch(function (error) {
    console.log(error);
  });
}






export async function footerText() {
  return Word.run(async (context) => {
    var sections = context.document.sections;
    context.load(sections, "body/style");
    return context.sync().then(function () {
      var section = sections.items[0];
      //

      var primaryFooter = section.getFooter("primary");
      var firstPageFooter = section.getFooter("firstPage");

      primaryFooter.insertText("welcome to AML", Word.InsertLocation.start);
      // primaryFooter.insertText("primaryFooter", Word.InsertLocation.start);
      // firstPageFooter.insertText("firstPageFooter", Word.InsertLocation.start);
      firstPageFooter.insertText("firstPageFooter", Word.InsertLocation.start);
      return context.sync().then(function () {
        // console.log("Completed");
      });
    });
  }).catch(function (error) {
    console.log(error);
  });
}

export async function footerImage() {
  return Word.run(async (context) => {
    var sections = context.document.sections;
    context.load(sections, "body/style");
    return context.sync().then(function () {
      var section = sections.items[0];
      //

      var primaryFooter = section.getFooter("primary");
      var firstPageFooter = section.getFooter("firstPage");

      primaryFooter.insertInlinePictureFromBase64("", Word.InsertLocation.start);
      primaryFooter.insertInlinePictureFromBase64("", Word.InsertLocation.start);
      // primaryFooter.insertText("primaryFooter", Word.InsertLocation.start);
      // firstPageFooter.insertText("firstPageFooter", Word.InsertLocation.start);
      //firstPageFooter.insertText("firstPageFooter", Word.InsertLocation.start);
      return context.sync().then(function () {
        console.log("Completed");
      });
    });
  }).catch(function (error) {
    console.log(error);
  });
}
