/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, window */

// The initialize function must be run each time a new page is loaded
Office.onReady(async () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;

  await Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  const ribbonJson = {
    actions: [
      {
        id: "executeWriteData",
        type: "ExecuteFunction",
        functionName: "writeData",
      },
    ],
    tabs: [
      {
        id: "CtxTab1",
        label: "Contoso Data",
        groups: [
          {
            id: "CustomGroup111",
            label: "Insertion",
            icon: [
              {
                size: 32,
                sourceLocation: "https://localhost:3000/assets/icon-32.png",
              },
              {
                size: 80,
                sourceLocation: "https://localhost:3000/assets/icon-80.png",
              },
            ],
            controls: [
              {
                type: "Button",
                id: "CtxBt112",
                actionId: "executeWriteData",
                enabled: false,
                label: "Write Data",
                superTip: {
                  title: "Data Insertion",
                  description: "Use this button to insert data into the document.",
                },
                icon: [
                  {
                    size: 32,
                    sourceLocation: "https://localhost:3000/assets/icon-32.png",
                  },
                  {
                    size: 80,
                    sourceLocation: "https://localhost:3000/assets/icon-80.png",
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };

  await Office.ribbon.requestCreateControls(ribbonJson);

  await new Promise((resolve) => {
    window.setTimeout(resolve, 2000);
  });

  await Office.ribbon.requestUpdate({
    tabs: [
      {
        id: "CtxTab1",
        visible: true,
      },
    ],
  });
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
