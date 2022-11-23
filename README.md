# Custom Functions

Use [this Google Workspace Add-on](https://workspace.google.com/marketplace/app/custom_functions/3868008326) in Google Sheets to instantly import custom functions, built using Apps Script.

# Table of contents
- [Installation](#installation)
    - [Usage](#usage)
- [Contribute](#contribute)
    - [Add new](#add-new)
        - [Must have](#must-have)
        - [Good to have](#good-to-have)
    - [Update existing](#update-existing)
- [License](#license-mit)

## Installation
1. Go to [the add-on page](https://workspace.google.com/marketplace/app/custom_functions/3868008326)
2. Click on Admin/individual install
3. Open [sheets.new](https://sheets.new)
4. Name the sheet (this is so the "Untitled Spreadsheet" gets saved to the Google Drive)
5. Enable your Google Apps Script API by navigating to [this link](https://script.google.com/home/usersettings)
    - Note: In case you have signed-in using multiple Google accounts in the same browser sessions, ensure that you're enabling the API with the same account in which the add-on has been installed
6. [Optional] Refresh the add-on
7. Click on any of the function you see from the grid and click to navigate to its page
8. Click on the IMPORT button

### Usage

Once you've imported a function, those would be as easy to use as a built-in function:

1. Click the cell where you want to use the function.
2. Type an equals sign (`=`) followed by the function name and any input value — for example, `=DOUBLE(A1)` — and press Enter.
3. The cell will momentarily display `Loading...`, then return the result.

## Contribute

In general, we'll follow the "fork-and-pull" Git workflow:

1. Fork this repo on GitHub
2. Work on your fork
    - All the files that you'll need to work with would be in the [functions](/functions/) folder
    - Make your additions or changes there
3. Commit changes to your own branch
4. Make sure you merge the latest from "upstream" and resolve conflicts (if any)
5. Push your work back up to your fork
6. Submit a [Pull request](https://github.com/custom-functions/google-sheets/pulls) so that we can review your changes

### Add new

Use the following as a template when adding a new function —

```
// @author WorkspaceDevs https://developers.google.com/apps-script/guides/sheets/functions#optimization
/**
 * Multiplies the input value by 2.
 *
 * @param {number} input The value or range of cells to multiply.
 * @return {number} The input multiplied by 2.
 * @customfunction
 */
function DOUBLE(input) {
  return Array.isArray(input) ?
    input.map(row => row.map(cell => cell * 2)) :
    input * 2;
}
```

#### Must have —

1. JSDoc styled comments
2. The description of the function needs to go all the way on the top (within JSDoc)
3. Function needs to have —
    - a clear/defined input `@param` (there can be more than 1 of these)
    - A single `@return` tag
4. JSDoc will need to end with the `@customfunction` tag
5. Ensure to rigorously test the function in your own Apps Script project

#### Good to have —

1. Have the `@author` tag added at the very first line of the file (if any)
2. While the function can return a single value, where possible, it would be good to be able to accomodate `Array` input and return too

### Update existing

Refer the **general** instructions layed out under the [Contribute](#contribute) section.

## License (MIT)

```
Copyright (c) 2022 Custom Functions

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```