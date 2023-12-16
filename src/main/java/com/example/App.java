package com.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App extends JFrame {
    /**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JTextField nameField, ageField, emailField;

    public App() {
        // Set up the JFrame
        super("User Details Form");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(450, 250);
        setLocationRelativeTo(null);

        // Create components
        JLabel nameLabel = new JLabel("Name:");
        JLabel ageLabel = new JLabel("Age:");
        JLabel emailLabel = new JLabel("Email:");

        nameField = new JTextField(30);
        ageField = new JTextField(5);
        emailField = new JTextField(30);

        JButton saveButton = new JButton("Save");
        saveButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                saveUserDetails();
            }
        });

        // Set up layout
        setLayout(new GridLayout(4, 2));

        // Add components to the JFrame
        add(nameLabel);
        add(nameField);
        add(ageLabel);
        add(ageField);
        add(emailLabel);
        add(emailField);
        add(new JLabel()); // Empty label for spacing
        add(saveButton);

        // Display the JFrame
        setVisible(true);
    }

    private void saveUserDetails() {
        String name = nameField.getText();
        String age = ageField.getText();
        String email = emailField.getText();

        if (name.isEmpty() || age.isEmpty() || email.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please fill in all fields.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        try {
            // Create a new Excel workbook
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("User Details");

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("Email");

            // Create data row
            Row dataRow = sheet.createRow(1);
            dataRow.createCell(0).setCellValue(name);
            dataRow.createCell(1).setCellValue(age);
            dataRow.createCell(2).setCellValue(email);

            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("UserDetails.xlsx")) {
                workbook.write(fileOut);
            }

            // Close the workbook
            workbook.close();

            JOptionPane.showMessageDialog(this, "User details saved successfully.", "Success", JOptionPane.INFORMATION_MESSAGE);

        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(this, "Error saving user details.", "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new App();
            }
        });
    }
}
