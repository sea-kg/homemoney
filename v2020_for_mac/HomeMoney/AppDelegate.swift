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
    var months: [PageMonth] = []

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
        months.removeAll()
        comboBoxMonthes.removeAllItems()
        let filePath = self.pathToFile.stringValue
        guard let file = XLSXFile(filepath: filePath) else {
            fatalError("XLSX file corrupted or does not exist")
        }
        if let workbooks: [Workbook] = try? file.parseWorkbooks() {
            if workbooks.count != 1 {
                fatalError("Expected one workbook")
            } else {
                let workbook: Workbook = workbooks[0]
                for (worksheetName, path) in try! file.parseWorksheetPathsAndNames(workbook: workbook) {
                    if let worksheetName = worksheetName {
                        if PageMonth.isPageMonth(name: worksheetName) {
                            print("This page is month")
                            months.append(PageMonth(name: worksheetName, path: path))
                        }
                        print("This worksheet has a name: \(worksheetName)")
                    }
                    
                    /*if let worksheet = try? file.parseWorksheet(at: path) {
                        for row in worksheet.data?.rows ?? [] {
                            for c in row.cells {
                                print(c)
                            }
                        }
                    }*/
                }
                months.sort { month1, month2 -> Bool in
                    return month1.name > month2.name
                }
                
                for m in months {
                    comboBoxMonthes.addItem(withObjectValue: m.name)
                }
                if months.count > 0 {
                    comboBoxMonthes.selectItem(at: 0)
                }
            }
        }
    }
}

