package com.excel.loader;

import com.gui.ControllerFrame;

public class SchedulerApp {
    /**
     * Main method for executing application.
     * 
     * @param args from command line
     */
    public static void main(String... args) {
        ExcelDataLoader dataloader = new ExcelDataLoader();
        new ControllerFrame(dataloader);
    }
}
