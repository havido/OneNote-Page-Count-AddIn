/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  /**
   * Insert your OneNote code here
   */
  try {
    await OneNote.run(async (context) => {

        // I want to count # of pages in the current section
        // Get the current page, section, and notebook.
        const page = context.application.getActivePage();
        const section = context.application.getActiveSection();
        const notebook = context.application.getActiveNotebook();
      
        section.pages.load("items");
        await context.sync();
        console.log("Number of pages in the current section: " + section.pages.items.length);

        // // Queue a command to add an outline to the page.
        // const html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, "Number of pages in the current section: " + section.pages.items.length);

        // Run the queued commands.
        // await context.sync();
    });
  } catch (error) {
      console.log("Error: " + error);
  }
}
