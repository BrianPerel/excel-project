package com.gui;

import java.awt.Button;
import java.awt.GridLayout;
import java.awt.Panel;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.logging.Logger;

import javax.swing.JFrame;
import javax.swing.WindowConstants;

import com.excel.loader.ExcelDataLoader;

public class ControllerFrame extends JFrame implements ActionListener {
    private static final long serialVersionUID = 3788675020358559959L;
    private final Button buttonStart = new Button("Start");
    private final Button buttonExit = new Button("Exit");
    private final Button switchToTextWindow = new Button("Text");

    private ExcelDataLoader dataLoader;
    private Panel mainPanel;

    private static final Logger logger_ = Logger.getLogger(ControllerFrame.class.getSimpleName());

    public ControllerFrame(ExcelDataLoader argDataLoader) {
        dataLoader = argDataLoader;
        mainPanel = new Panel();

        setLayout(new GridLayout(2, 1));
        setSize(500, 300);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent argEvent) {
                buttonExit.doLayout();
            }
        });

        buttonStart.addActionListener(this);
        mainPanel.add(buttonStart);
        buttonExit.addActionListener(this);
        mainPanel.add(buttonExit);

        switchToTextWindow.addActionListener(this);
        mainPanel.add(switchToTextWindow);

        add(mainPanel);
        setVisible(true);
    }

    @Override
    public void actionPerformed(ActionEvent argEvent) {
        if (argEvent.getSource() == buttonStart) {
            dataLoader.initializeProcess();
            System.exit(9);
        }

        if (argEvent.getSource() == buttonExit) {
            System.exit(9);
        }
        if (argEvent.getSource() == switchToTextWindow) {
            textButtonPressed(argEvent);
        }
    }

    private void textButtonPressed(ActionEvent argEvent) {
        this.dispose();
        TextFrame textFrame = new TextFrame(argEvent);
        textFrame.setVisible(true);
    }

}
