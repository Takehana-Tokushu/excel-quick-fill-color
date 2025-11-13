/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, Excel */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

// Color definitions - you can easily change these values
const COLORS = {
  YELLOW: "#FFFF00",
  ORANGE: "#FFA500",
  GRAY: "#A9A9A9"
};

/**
 * Fill selected cells with a specific color
 * @param {string} color - The hex color code
 * @param {Office.AddinCommands.Event} event
 */
async function fillCellsWithColor(color, event) {
  try {
    await Excel.run(async (context) => {
      // Get all selected ranges (supports non-contiguous selection with Ctrl)
      const selectedRanges = context.workbook.getSelectedRanges();
      
      // Load the areas to get each individual range
      selectedRanges.load("areas");
      
      await context.sync();
      
      // Apply color to each range in the selection
      const areas = selectedRanges.areas;
      areas.load("items");
      
      await context.sync();
      
      // Loop through each area and apply the color
      for (let i = 0; i < areas.items.length; i++) {
        areas.items[i].format.fill.color = color;
      }
      
      // Sync all changes to Excel
      await context.sync();
      
      console.log(`Successfully filled ${areas.items.length} range(s) with color: ${color}`);
    });
  } catch (error) {
    console.error("Error filling cells:", error);
    
    // Show error message to user
    if (error && error.message) {
      showNotification("Error", "Failed to fill cells: " + error.message);
    }
  } finally {
    // IMPORTANT: Always complete the event
    if (event && event.completed) {
      event.completed();
    }
  }
}

/**
 * Clear fill color from selected cells
 * @param {Office.AddinCommands.Event} event
 */
async function clearCellsFill(event) {
  try {
    await Excel.run(async (context) => {
      // Get all selected ranges (supports non-contiguous selection with Ctrl)
      const selectedRanges = context.workbook.getSelectedRanges();
      
      // Load the areas to get each individual range
      selectedRanges.load("areas");
      
      await context.sync();
      
      // Get each area
      const areas = selectedRanges.areas;
      areas.load("items");
      
      await context.sync();
      
      // Loop through each area and clear the fill
      for (let i = 0; i < areas.items.length; i++) {
        areas.items[i].format.fill.clear();
      }
      
      // Sync all changes to Excel
      await context.sync();
      
      console.log(`Successfully cleared fill from ${areas.items.length} range(s)`);
    });
  } catch (error) {
    console.error("Error clearing fill:", error);
    
    // If the clear() method doesn't work, try alternative approach
    try {
      await Excel.run(async (context) => {
        const selectedRanges = context.workbook.getSelectedRanges();
        selectedRanges.load("areas");
        await context.sync();
        
        const areas = selectedRanges.areas;
        areas.load("items");
        await context.sync();
        
        // Try setting pattern to none for each area
        for (let i = 0; i < areas.items.length; i++) {
          areas.items[i].format.fill.pattern = Excel.FillPattern.none;
        }
        
        await context.sync();
        console.log("Cleared fill using pattern method");
      });
    } catch (fallbackError) {
      console.error("Fallback method also failed:", fallbackError);
      showNotification("Error", "Failed to clear fill: " + error.message);
    }
  } finally {
    // IMPORTANT: Always complete the event
    if (event && event.completed) {
      event.completed();
    }
  }
}

/**
 * Fill selected cells with yellow
 * @param {Office.AddinCommands.Event} event
 */
function fillYellow(event) {
  fillCellsWithColor(COLORS.YELLOW, event);
}

/**
 * Fill selected cells with orange
 * @param {Office.AddinCommands.Event} event
 */
function fillOrange(event) {
  fillCellsWithColor(COLORS.ORANGE, event);
}

/**
 * Fill selected cells with gray
 * @param {Office.AddinCommands.Event} event
 */
function fillGray(event) {
  fillCellsWithColor(COLORS.GRAY, event);
}

/**
 * Clear fill color from selected cells
 * @param {Office.AddinCommands.Event} event
 */
function clearFill(event) {
  clearCellsFill(event);
}

/**
 * Show notification to user (helper function)
 * @param {string} header - Notification header
 * @param {string} message - Notification message
 */
function showNotification(header, message) {
  // Only show notifications if Office.context.ui is available
  if (Office.context.ui && Office.context.ui.displayDialogAsync) {
    console.log(header + ": " + message);
  }
}

// Register the functions with Office
// These must match the FunctionName values in manifest.xml
if (Office.actions) {
  Office.actions.associate("fillYellow", fillYellow);
  Office.actions.associate("fillOrange", fillOrange);
  Office.actions.associate("fillGray", fillGray);
  Office.actions.associate("clearFill", clearFill);
}
