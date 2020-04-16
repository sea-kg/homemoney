//
//  AppDelegate.swift
//  HomeMoney
//
//  Created by sea-kg on 01.04.2020.
//  Copyright © 2020 sea-kg. All rights reserved.
//

import Cocoa
// import CoreXLSX


@NSApplicationMain
class AppDelegate: NSObject, NSApplicationDelegate, NSOpenSavePanelDelegate {

    @IBOutlet weak var window: NSWindow!


    func applicationDidFinishLaunching(_ aNotification: Notification) {
        // Insert code here to initialize your application
    }

    func applicationWillTerminate(_ aNotification: Notification) {
        // Insert code here to tear down your application
    }
    @IBOutlet weak var pathToFile: NSTextField!
    @IBOutlet weak var comboBoxMonthes: NSComboBox!
    
    @IBAction func selectFile(_ sender: NSButton) {
        let openPanel = NSOpenPanel();
        openPanel.title = "Выберите файл xlsx"
        openPanel.message = ""
        openPanel.showsResizeIndicator = true;
        openPanel.canChooseDirectories = false;
        openPanel.canChooseFiles = true;
        openPanel.allowsMultipleSelection = false;
        openPanel.canCreateDirectories = false;
        openPanel.delegate = self;
        
        openPanel.begin { (result) -> Void in
            if (result == NSApplication.ModalResponse.OK) {
                let path = openPanel.url!.path
                print("path: \(path)");
                self.pathToFile.stringValue = path
                self.loadMonthesFromFile()
            }
        }
    }
    
    func loadMonthesFromFile() {
        let filePath = self.pathToFile.stringValue
        guard let file = XLSXFile(filepath: filePath) else {
            fatalError("XLSX file corrupted or does not exist")
        }
        if let paths = try? file.parseWorksheetPaths() {
            for path in paths {
                if let worksheet = try? file.parseWorksheet(at: path) {
                    print(worksheet.properties ?? nil)
                    
                    // print(worksheet)
                }
                // worksheet
                /*for row in worksheet.data?.rows ?? [] {
                    for c in row.cells {
                        print(c)
                    }
                }*/
            }
        }
    }
}

