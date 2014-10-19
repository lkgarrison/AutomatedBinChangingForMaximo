/* 
 * This program will help automate the process of changing bin locations in Maximo v4
 * It is used along with an excel document that contains the item numbers in one column and the
 * correct bin location corresponding to each item in the column to the right of the item numbers.
 * 
 * Use with the accompanying documentation for specific instructions with accompanying screenshots
 * 
 * New in this version: has graphical features (for extra safety), closes Maximo when finished changing bins
 * 
 */

import java.awt.Robot;
import java.awt.datatransfer.*;
import java.awt.event.*;
import java.awt.Toolkit;
import java.awt.Dimension;
import java.io.IOException;
import javax.swing.*;

public class ChangeBinLocationExitMaximo implements ActionListener {
	
	public static void main(String[] args) throws UnsupportedFlavorException, IOException {
		init();
		changeBinsInMaximo();
	}
	
	private static void changeBinsInMaximo() throws UnsupportedFlavorException, IOException {		
		
		String itemNum;								// initialize
		boolean itemNumIsValid;
		
		// while there is a valid item number to change the bin:
		while(true) {
			itemNum = getNextItem();
			itemNumIsValid = isValidItemNum(itemNum);
			if (itemNumIsValid == !true) break;		// this indicates the program has finished changing all designated items in the list
			
			pasteNextItem();						// enters copied item # in Maximo
			checkForDuplicates();					// will not change bins if item has duplicate bins
			if (isDuplicate == true) {
				addItemStatusToExcel();
				continue;
			}
			adjustPhysicalCount("0"); 				// set physical count to 0
			pasteNewBin();							// set new default bin
			recheckItemNum();
			adjustPhysicalCount("current balance");	// set physical count to previous bin's current balance
			reconcileBalance();
			deleteOldBinLine();
			addItemStatusToExcel();
		}
	}
	
	private static void displayFinishedMessage() {
		userDialog(completeMessage, "Bin Changes Complete");	// dialog box to tell user when the operation is finished
		bot.waitForIdle();	// this will wait for event queue to empty (ie, when the user closes the "Bin changes complete" window) before shutting down the program
							// this way, the message window will stay open and then the program shuts down 100% as soon as the user clicks OK on the message
	}
	
	private static void exitAll() {
		pointAndClick(maximoX, maximoY);		// click on a safe place in Maximo window
		bot.delay(300);
		pointAndClick(fileX, fileY);
		bot.delay(1000);
		pointAndClick(exitAllX + fileExtraX, exitAllY);
	}
		
	private static void createStopButton() {	
		userDialog(startConfirmationMessage, "Start Confirmation");
		
		// opens up the stop button, which will close if the mouse is moved over it
		stopFrame = new stopFrame(screenWidth, screenHeight);
	
		bot.delay(3000);	// gives user 3 seconds to let go of mouse
				
		// this code should be removed after testing is finished. not necessary for the user. and set speedConstant = 1 @ bottom.
		/* testing purposes only:
		String speedConstantString = JOptionPane.showInputDialog("Speed Multiplier: ");
		speedConstant = Integer.parseInt(speedConstantString);
		*/
	}
	
	/* this method checks to see if the item # has duplicate locations. If it does, then 
	 * this item will be skipped an and error message will be provided next to the item in the excel sheet
	 */
	private static void checkForDuplicates() throws UnsupportedFlavorException, IOException {
		isDuplicate = true;		// assume they are duplicates until proven otherwise
		
		// gets default bin's current balance		
		copyCurrentBalance();
		
		// gets total current balance
		pointAndClick(totalCurBalX, totalCurBalY);
		bot.mousePress(mask);
		bot.mouseMove((int) (totalCurBalX + selectionExtraX), (int) totalCurBalY);
		bot.mouseRelease(mask);
		bot.delay(speedConstant * 50);
		String totalCurBal = copy();
		bot.delay(50 * speedConstant);

		if(oldBinCurBal.equals(totalCurBal)) isDuplicate = false;
	}
	
	private static void recheckItemNum() throws UnsupportedFlavorException, IOException {
		String itemNum = copyItemNum();
		checkForValidItemNum(itemNum);
	}
	
	private static void checkForValidItemNum(String itemNum) throws UnsupportedFlavorException, IOException {
		itemNum = itemNum.replaceAll("\\n", "");	// removes the newline (enter) character which is added by copying from excel
	
		boolean isValidItemNum = isValidItemNum(itemNum);
		if (isValidItemNum == false)  {				
			exitAll();
			displayFinishedMessage();
			System.exit(0);							// end program if there is an invalid item #
		}
	}
	
	private static boolean isValidItemNum(String itemNum) {
			boolean isValidItemNum = false;
			
			if (itemNum.length() != 10) return isValidItemNum;
			else if (itemNum.charAt(3) != '-') return isValidItemNum;
			for (int i = 0; i < 3; i++) {	// checks characters before the item num's "-"
				if (itemNum.charAt(i) < '0' || itemNum.charAt(i) > '9') return isValidItemNum;	  // if character isn't a number
			}
			for (int i = 4; i < itemNum.length(); i++) {	// checks from the "-" onward 
				if (itemNum.charAt(i) < '0' || itemNum.charAt(i) > '9') return isValidItemNum;	  // if character isn't a number
			}
			
			isValidItemNum = true;
			return isValidItemNum;
	}
	
	private static String copyItemNum() throws UnsupportedFlavorException, IOException {
		// copies pasted item number field in Maximo and returns this selection
		pointAndClick(itemNumSelectionX, itemNumSelectionY);
		bot.delay(speedConstant * 75);
		bot.mousePress(mask);
		bot.mouseMove((int) itemNumSelectionEndX, (int) itemNumSelectionY);
		bot.mouseRelease(mask);
		String itemNum = copy();
		return itemNum;
	}
	
	private static void addItemStatusToExcel() {
		// sets current cell to be the "success" column in excel
		pointAndClick(excelX, excelY);
		bot.delay(speedConstant * 25);
		moveRight();
		
		if (isDuplicate == false) {
			addY();			// adds Y to "successfully added" column
		} else {
			moveRight(); 
			addNO();		// adds "NO" to "Success?" Column
			moveRight();
			StringSelection duplicateError = new StringSelection("Duplicate Bins. Item Skipped. Please review this item's bin locations.");
			clipboard.setContents(duplicateError, duplicateError);
			paste();	// will paste duplicate bin error message
			moveLeft();
		}
		
		enter();			// gets ready to repeat the whole procedure. gets in starting position (next item #)
		moveLeft();  	
		moveLeft();
		moveLeft();
	}
	
	private static void deleteOldBinLine() {		// deletes the line entry of the previous, zeroed out bin
		pointAndClick(oldBinLineX, oldBinLineY);
		bot.delay(speedConstant * 50);
		delete();
		save();
	}
	
	private static void reconcileBalance() {		// clicks "reconcile balance" in Maximo
		pointAndClick(actionX, actionY);
		bot.delay(speedConstant * 400);		
		bot.mouseMove((int) inventoryAdjustmentsX, (int) inventoryAdjustmentsY);
		bot.delay(speedConstant * 150);
		bot.mouseMove((int) inventoryAdjustmentsX2, (int) inventoryAdjustmentsY);	// must have this horizontal slide to get next menu to come up
		bot.delay(speedConstant * 800);
		bot.mouseMove((int) reconcileBalancesX, (int) reconcileBalancesY);
		bot.delay(speedConstant * 100);
		mouseClick();
		enter();
	}
	
	private static void pasteNewBin() throws UnsupportedFlavorException, IOException {
		getNextBin();
		pointAndClick(defaultBinX, defaultBinY);
		paste();
		save();
	}
	
	private static void getNextBin() throws UnsupportedFlavorException, IOException {
		pointAndClick(excelX, excelY);
		moveRight();
		bot.delay(speedConstant * 50);
		copy();
		pointAndClick(maximoX, maximoY);	// bring maximo to foreground again
	}
	
	private static void adjustPhysicalCount(String key) {
		pointAndClick(actionX, actionY);
		bot.delay(speedConstant * 400);		
		bot.mouseMove((int) inventoryAdjustmentsX, (int) inventoryAdjustmentsY);
		bot.delay(speedConstant * 150);
		bot.mouseMove((int) inventoryAdjustmentsX2, (int) inventoryAdjustmentsY);	// must have this horizontal slide to get next menu to come up
		bot.delay(speedConstant * 850);
		bot.mouseMove((int) reconcileBalancesX, (int) reconcileBalancesY);
		bot.delay(speedConstant * 200);
		pointAndClick(adjustPhysicalCountX, adjustPhysicalCountY);
		bot.delay(speedConstant * 300);
		if (key.equals("0")) {		// set physical count to 0
			bot.keyPress(KeyEvent.VK_0);
		} else {
			// manually set clipboard:
			StringSelection actualBal = new StringSelection(oldBinCurBal);
			clipboard.setContents(actualBal, actualBal);
			paste();		
		}
		enter();
	}
	
	private static void copyCurrentBalance() throws UnsupportedFlavorException, IOException {
		/* click mouse CurBal box. Then press, move to the right, and release to select that current balance and copy it */
		pointAndClick(oldCurBalBeforeNewX, oldCurBalBeforeNewY);
		bot.mousePress(mask);
		bot.mouseMove((int) (oldCurBalBeforeNewX + selectionExtraX), (int) oldCurBalBeforeNewY);
		bot.mouseRelease(mask);
		bot.delay(speedConstant * 50);
		oldBinCurBal = copy();
	}
	
	private static String getNextItem() throws UnsupportedFlavorException, IOException {
		// will go to Excel and copy the next item to the clip board
		pointAndClick(excelX, excelY);
		String copiedSelection = copy();
		
		// If the copied selection is not a valid item number, the program will  end
		copiedSelection = copiedSelection.replaceAll("\\n", "");	// removes the newline (enter) character which is added by copying from excel
		checkForValidItemNum(copiedSelection);
		return copiedSelection;
	}
	
	private static void pasteNextItem () throws UnsupportedFlavorException, IOException {
		// will paste the Maximo Item # without clicking directly into the item # field, go back to excel and press ESC (removes item #
		// from clipboard), and then go back to Maximo and vefity it is a valid item # before pressing enter.
		pointAndClick(maximoX, maximoY);
		bot.delay(speedConstant * 100);
		esc();	// closes out any item that is already on the screen
		paste();
		
		// remove item # from clipboard
		pointAndClick(excelX, excelY);
		esc();
		pointAndClick(maximoX, maximoY);
		
		String itemNum = copyItemNum();
		checkForValidItemNum(itemNum);	// will ensure the item number that was pasted is valid before proceeding
		enter();
	}
	
	private static void userDialog(String message, String title) {
		JFrame userDialogFrame = new JFrame(title);
		
		userDialogFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		JOptionPane.showMessageDialog(userDialogFrame, 
				message,
				title,
				JOptionPane.INFORMATION_MESSAGE);
	}
	
	private static void mouseClick () {
		if (stopFrame.isTerminated == true) {
			bot.delay(10000);
			System.exit(0);
		}
		bot.mousePress(mask);
		bot.mouseRelease(mask);
		bot.delay(speedConstant * 200);
	}
	
	private static void pointAndClick (double x, double y) {
		if (stopFrame.isTerminated == true) {
			bot.delay(10000);
			System.exit(0);
		}
		bot.mouseMove( (int) x, (int) y);
		mouseClick();
	}
	
	private static String copy() throws UnsupportedFlavorException, IOException {
		bot.delay(speedConstant * 50);		// without this, sometimes it clicks away before it copies
		bot.keyPress(KeyEvent.VK_CONTROL);
		bot.keyPress(KeyEvent.VK_C);
		bot.delay(speedConstant * 50);
		bot.keyRelease(KeyEvent.VK_C);
		bot.keyRelease(KeyEvent.VK_CONTROL);
		bot.delay(speedConstant * 200);
		String copiedSelection = (String) clipboard.getData(DataFlavor.stringFlavor);
		return copiedSelection;
	}

	private static void paste() {
		bot.delay(speedConstant * 50);
		bot.keyPress(KeyEvent.VK_CONTROL);
		bot.keyPress(KeyEvent.VK_V);
		bot.keyRelease(KeyEvent.VK_V);
		bot.keyRelease(KeyEvent.VK_CONTROL);
		bot.delay(speedConstant * 200);
	}
	
	private static void save() {
		// save keyboard shortcut ctrl + S
		bot.keyPress(KeyEvent.VK_CONTROL);
		bot.keyPress(KeyEvent.VK_S);
		bot.keyRelease(KeyEvent.VK_S);
		bot.keyRelease(KeyEvent.VK_CONTROL);
		bot.delay(speedConstant * 50);
	}
	
	private static void delete() {
		// delete keyboard shortcut ctrl + d
		bot.keyPress(KeyEvent.VK_CONTROL);
		bot.keyPress(KeyEvent.VK_D);
		bot.keyRelease(KeyEvent.VK_D);
		bot.keyRelease(KeyEvent.VK_CONTROL);
		bot.delay(speedConstant * 50);
	}
	
	private static void enter() {
		// will simulate an enter keypress and release
		bot.keyPress(KeyEvent.VK_ENTER);
		bot.keyRelease(KeyEvent.VK_ENTER);
		bot.delay(speedConstant * 50);
	}
	
	private static void esc() {
		// simulates pressing escape key
		bot.keyPress(KeyEvent.VK_ESCAPE);
		bot.keyRelease(KeyEvent.VK_ESCAPE);
		bot.delay(speedConstant * 50);
	}
	
	private static void moveRight() {
		bot.keyPress(KeyEvent.VK_RIGHT);
		bot.keyRelease(KeyEvent.VK_RIGHT);
		bot.delay(speedConstant * 50);
	}
	
	private static void moveLeft() {
		bot.keyPress(KeyEvent.VK_LEFT);
		bot.keyRelease(KeyEvent.VK_LEFT);
		bot.delay(speedConstant * 50);
	}
	
	private static void addY() {
		// will type X
		bot.keyPress(KeyEvent.VK_Y);
		bot.keyRelease(KeyEvent.VK_Y);
		bot.delay(speedConstant * 50);
	}
	
	private static void addNO() {
		// will type X
		bot.keyPress(KeyEvent.VK_N);
		bot.keyRelease(KeyEvent.VK_N);
		bot.delay(speedConstant * 50);
		bot.keyPress(KeyEvent.VK_O);
		bot.keyRelease(KeyEvent.VK_O);
		bot.delay(speedConstant * 50);
	}
	
	private static void init() {
		// initialize robot:
		try {
			bot = new Robot();
		} catch (Exception failed) {
			System.err.println("Failed instantiating Robot: " + failed);
		}
		
		mask = InputEvent.BUTTON1_DOWN_MASK;
		//	bot.delay(2000); 	// setup time
		
		createStopButton();
	}
	
	private static Toolkit tk = Toolkit.getDefaultToolkit();
	private static Dimension dimension = tk.getScreenSize();
	private static int screenWidth = dimension.width;
	private static int screenHeight = dimension.height;
	private static Clipboard clipboard = tk.getSystemClipboard();
	
	private static Robot bot = null;
	private static int mask;
	private static stopFrame stopFrame;
	
	private static boolean isDuplicate;
	private static String oldBinCurBal;
	private static int speedConstant = 1;		// ability to slow program down to watch more carefully
	
	/* below are constants used to accurately interact with Maximo */
	private static final double actionX = .122 * screenWidth;		// location of "Action" menu
	private static final double actionY =  .033 * screenHeight;
	private static final double inventoryAdjustmentsX = actionX;	// location of "Inventory Adjustments" inside action menu
	private static final double inventoryAdjustmentsX2 = inventoryAdjustmentsX + .04375 * screenWidth;
	private static final double inventoryAdjustmentsY =  .238 * screenHeight;
	private static final double reconcileBalancesX = .23875 * screenWidth;		// location of "Reconcile Balances" inside "Inventory Adjustments"
	private static final double reconcileBalancesY = inventoryAdjustmentsY;
	private static final double adjustPhysicalCountX = reconcileBalancesX;		// location of "Adjust Physical Count"
	private static final double adjustPhysicalCountY = .31667 * screenHeight;
	private static final double physicalCountX = .2369 * screenWidth;	// location of physical balance box
	private static final double physicalCountY = .411 * screenHeight;
	private static final double oldCurBalBeforeNewX = .193125 * screenWidth;
	private static final double oldCurBalBeforeNewY = .411111 * screenHeight;
	private static final double selectionExtraX = .06625 * screenWidth; // necessary to move mouse over this far while pressed down to select Physical Count number
	private static final double defaultBinX = .45625 * screenWidth;		// location of Default Bin box
	private static final double defaultBinY = .15889 * screenHeight;
	private static final double itemNumSelectionX = .1625 * screenWidth;	// location of Maximo item # box. Also used to copy the number in that text field
	private static final double itemNumSelectionY = .1375 * screenHeight;
	private static final double itemNumSelectionEndX = .1025 * screenWidth;
	private static final double oldBinLineX = .073125 * screenWidth;
	private static final double oldBinLineY = .434444 * screenHeight;
	private static final double totalCurBalX = .186875 * screenWidth;	// location of total current balance of part (across all bin locations)
	private static final double totalCurBalY = .21444 * screenHeight;
	private static final double fileX = .06125 * screenWidth;				// location of text menu "File"
	private static final double fileY = .03778 * screenHeight;
	private static final double exitInvCtrlX = fileX;			// location of file/"Exit Inventory Control"
	private static final double exitInvCtrlY = .23889 * screenHeight;
	private static final double exitAllX = fileX;						// location of file/"Exit All"
	private static final double exitAllY = .261111 * screenHeight;
	private static final double fileExtraX = .025 * screenWidth;	// must move over horizontally by this amount for it to recognize the selection in the (file) text menu
	private static final double excelX = .90875 * screenWidth;		// safe place to click mouse to bring Excel to foreground again
	private static final double excelY = .031111 * screenHeight;
	private static final double maximoX = .020625 * screenWidth;	// safe place to click to bring Maximo to foreground. Shouldn't trigger any of the text menus
	private static final double maximoY = .02 * screenHeight;
	
	// the messages that are displayed in the dialog boxes are written out here:
	private static final String completeMessage = "All designated items now have the correct bin locations in Maximo. Please review the Excel spreadsheet to ensure the program did not terminate early due to an invalid item #. Press OK.";
	private static final String startConfirmationMessage = "This program will move your mouse automatically to change the default bin locations of the specified Maximo item numbers. Move your mouse over the STOP button at any time to end the program. Click OK to continue and remove your hand from the mouse.";
	
	public void actionPerformed(ActionEvent e) {
		String cmd = e.getActionCommand();
		if (cmd.equals("STOP")) {
			System.exit(0);
		}
	}
	
}
