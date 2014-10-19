/* This class is to be used with the ChangeBinLocation program. It provides the framework to get 
 * the window that will kill the program if the mouse is hovered over it, a vital safety feature.
 */

import java.awt.event.*;
import javax.swing.*;

public class stopFrame extends MouseAdapter {

	public stopFrame(int screenWidth, int screenHeight){
		
		double frameStartX = 0;
		double frameStartY = .65 * screenHeight;
		double frameWidth = .5 * screenWidth;
		double frameHeight = screenHeight - frameStartY;
		
		
		JFrame stopFrame = new JFrame();
		stopFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		stopFrame.setSize((int) frameWidth, (int) frameHeight);
		stopFrame.setLocation((int) frameStartX, (int) frameStartY);
		stopFrame.setTitle("Stop Button");
		stopFrame.setVisible(true);
		
		stopFrame.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e) {
			System.exit(0);
		}
		});
		
		stopFrame.addMouseMotionListener(new MouseAdapter() {
			public void mouseMoved(MouseEvent e) {
				System.exit(0);
			}
		});
	
		JButton stopButton = new JButton("STOP");
		stopButton.setSize(20, 20);
		stopButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				terminateProgram();
				System.exit(0);
			}
		});
		stopFrame.add(stopButton);
		
		
		stopButton.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e) {
			terminateProgram();
			System.exit(0);
		}
		});
		
		stopButton.addMouseMotionListener(new MouseAdapter() {
			public void mouseMoved(MouseEvent e) {
				terminateProgram();
				System.exit(0);
			}
		});
		
	}
	
	public static void terminateProgram() {
		isTerminated = true;
		JFrame terminateProgramFrame = new JFrame("Terminate Program");
		terminateProgramFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		JOptionPane.showMessageDialog(terminateProgramFrame, 
				terminateProgramMessage,
				"Terminate Program",
				JOptionPane.INFORMATION_MESSAGE);
	}
	

	public static boolean isTerminated = false;
	private static final String terminateProgramMessage = "The program has been terminated. Ensure that the most recent Maximo item number has not been only partially changed. Click OK to continue.";
	
}
