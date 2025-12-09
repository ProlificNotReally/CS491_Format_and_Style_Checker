// Copyright 2025 oh826
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     https://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


/* src/taskpane/components/fixers.js */

export const fixBasicFormatting = async () => {
  await Word.run(async (context) => {
    // 1. Fix Font Name (Global change to Times New Roman)
    context.document.body.font.name = "Times New Roman";

    // 2. Fix Font Size (Global normalization to 12pt)
    // Note: This enforces 12pt everywhere.
    context.document.body.font.size = 12;

    // 3. Fix Paper Size (Must iterate through sections)
    const sections = context.document.sections;
    context.load(sections, "pageSetup");
    await context.sync();

    sections.items.forEach((section) => {
      // Sets paper size to 8.5x11 Letter
      section.pageSetup.paperSize = "Letter"; 
    });

    await context.sync();
    console.log("Basic formatting fixed: Font, Size, and Paper Layout.");
  });
};

