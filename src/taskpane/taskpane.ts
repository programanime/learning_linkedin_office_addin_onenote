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
  await OneNote.run(async context => {
    let notebook = context.application.getActiveNotebook();
    let page = context.application.getActivePage();
    page.addOutline(40,90, '<p>hey hello</p>');
    context.load(notebook);
    return context.sync().then(() => {
      document.getElementById("mensaje").innerHTML = notebook.name;
    });
  })
}
