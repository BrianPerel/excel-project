package com.gui;

import java.awt.Button;
import java.awt.GridLayout;
import java.awt.TextField;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.logging.Logger;

import javax.swing.JFrame;
import javax.swing.WindowConstants;

public class TextFrame extends JFrame implements ActionListener {
    private static final long serialVersionUID = -8859273687713437163L;
    private TextField textField;
    private final Button buttonText = new Button("Save Text");

    private static final Logger logger_ = Logger.getLogger(TextFrame.class.getSimpleName());

    public TextFrame(ActionEvent argEvent) {
        textField = new TextField(20);

        setLayout(new GridLayout(2, 1));
        setSize(500, 300);
        setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        buttonText.addActionListener(this);
        add(buttonText);
        add(textField);
    }

    @Override
    public void actionPerformed(ActionEvent argEvent) {
        if (argEvent.getSource() == buttonText) {
            logger_.info(textField.getText());
            textField.setText("");
        }
    }
}
