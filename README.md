# Automating content development using Google Apps Script


## Overview

Content developers use applications such as Google Slides and Microsoft Powerpoint to create presentations. 

In Google Slides, there is a feature called Google Apps Script which enables you to write custom functions that perform manual editing tasks in code. Almost any manual action that you can perform in Google Slides can be automated.

In this tutorial, you will add functions and menu items that enable developers to apply multiple style changes to selected text in one step. Automating frequently used tasks such as these saves time and promotes consistency.

This guide assumes only basic familiarity with Google Slides and an object-oriented programming language such as Javascript. 



## Getting Started

In this tutorial, you will add the following function that applies multiple style changes to selected text in a presentation.

![sample code](./images/01-sample-code.png)

Lines 5 and 6 get the selected text; lines 10 to 13 apply style changes. 

Then you will add two wrapper functions that call this function with predetermined style changes, and add two menu items that users can select to run these functions as shown below.

![slides menu](./images/28-slides-menu.png)


## Create a Gmail account

To complete this lesson, you need a temporary Gmail account. 

Google Slides comes with Gmail.  However, running scripts requires two-factor authentication disabled and access to personal data, so it is safer and faster to run these steps from a temporary account.

Open a web browser.

Go to [gmail.com](https://gmail.google.com/) and create a Gmail account; it only takes about 60 seconds.

When you finish creating the account, Gmail opens.

## Create a new presentation in Google Slides

In the main Gmail page, click the **Google apps** icon (top-right.)

![Google apps](./images/02-gapps-selector.png)

In the popup dialog, select **Drive**.  

> Google Drive is where you can create Google Slides.

Click the **+ New** button (top-left) and select **Google Slides** to create a new presentation.

## Add some text to format

You need some text in slides to select, so you might as well add that now.

In the slide editor (center of page), click the title and subtitle text boxes and enter some text that you can format later.

![enter text](./images/04-enter-text2.png)

## Add functions that format text

Next, you define functions that will apply style changes when users select text to change.

Click **Extensions** in the main menu and select **Apps Script** to open the script editor. 

> **NOTE:** If Extensions is not listed, widen the browser window until it shows.

![Open Apps Script](./images/05-open-apps-script.png)

Copy the code below and paste it in the main editor window so it replaces the default, empty function.

```
function setRedConsolas() {
   setColoredFont("#FF0000", "Consolas");
}

function setGreenConsolasItalic() {
   setColoredFont("#00b050", "Consolas", true);
}

function setColoredFont(color, font, italic=false) {
  var color, font, italic;
  
  // get text selection
  var selection = SlidesApp.getActivePresentation().getSelection();
  var text = selection.getTextRange();

  // format text
  if(text.getLength() > 0) {
    var style = text.getTextStyle();
    style.setForegroundColor(color);
    style.setFontFamily(font);
    style.setItalic(italic);
  }
}
```

Here is what a successful replacement looks like.

![succesful paste](./images/06a-successful-paste.png)

There are three functions: two functions (**setRedConsolas** and **setGreenConsolasItalic**) specify different style changes and call the third function (**setColoredFont**) to apply changes.

Click the **Save Project** icon to save your functions before running and testing.

![save icon](./images/07-save-icon.png)

## Test and authorize your functions

In the next section, you add custom menu items that expose these functions to end users. But first you want to test and authorize them to act on data.

Return to slides and select some text to highlight (a word or two is all you need).

![select text](./images/06b-select-text.png)

Leave the text selected and return to script editor.

Select **setRedConsolas** from the function drop-down selector (if not selected).

![select red](./images/08-select-red.png)

Click **Run** to call **setRedConsolas**. This tests the function and requests the necessary privileges to access and modify text in presentations.

![click run](./images/09a-click-red.png)

Click **Review permissions**.

![review perms](./images/20-review-perms.png)

Select your temporary account.

![click account](./images/21-choose-account.png)

Click **Advanced**.

![click advanced](./images/22-click-advanced.png)

Click **Go to untitled project (unsafe)**. This is just a warning.

![click proj](./images/23-click-go-proj.png)

Click **Allow**.

![click allow](./images/24-click-allow.png)

You should see the following messages in the Execution log appear as shown below.

> **NOTE:** You may need to run the script a second time after authorizing it.

![no errors](./images/25-no-errors.png)

If you get an error like the one below, no text was selected. Reselect some text in slides and rerun the function.

![script error](./images/26-script-error-len0.png)

Return to slides and notice the styles applied.

Select some other text and test **setGreenConsolasItalic**.

Congratulations! Your functions are working.


## Add functionality to the main menu

Now it's time to expose these functions to end users.

Copy and paste the following **onOpen** function into the top of the script editor.

```
function onOpen() {
  var ui = SlidesApp.getUi();
  
  ui.createMenu('Training Tools')
      .addItem('red consolas', 'setRedConsolas')
      .addItem('green consolas italic', 'setGreenConsolasItalic')
      .addToUi();
}
```

> **NOTE:** This function runs when a presentation containing it is opened for editing. It adds a menu item with two drop-down items. Each drop-down calls a function and applies style changes.

Save the changes to the script. 

Close and reopen the presentation.

> **NOTE:** You will be required to reauthorize the script so it can modify the main menu.

After the presentation opens, note the added menu item at the end of the main menu. 

![menu item](./images/34-slides-menu.png)

> **NOTE:** If you do not see the added menu item, widen the browser window.

Select some text in the slide edit window and select one of the drop-down menu items.

Congratulations! You are now using custom functions to save time and promote consistency in your content.

## Doing more with Google Docs

If there was more time, you could learn how to customize Google Docs to insert custom quoted text boxes and code blocks like those found in Markdown.

Here is what a custom menu looks like in Google Docs with two functions that insert the preformatted text boxes shown (text is added by users after boxes are added).

![text boxes](./images/35-docs-menu.png)


