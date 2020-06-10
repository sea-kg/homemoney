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
    var classifications: PageClassifications? = nil
    var defaultTextColor: NSColor? = nil
    func applicationDidFinishLaunching(_ aNotification: Notification) {
        // Insert code here to initialize your application
        let text = self.LogProcessing.string
        let range = (text as NSString).range(of: text)
        
        let color: NSColor = self.LogProcessing.attributedString().attribute(NSAttributedString.Key.foregroundColor, at: 0, effectiveRange: nil) as! NSColor
        print("color = \(color.cgColor)")
        self.defaultTextColor = self.LogProcessing.textColor
    }

    func applicationWillTerminate(_ aNotification: Notification) {
        // Insert code here to tear down your application
    }
    @IBOutlet weak var pathToFile: NSTextField!
    @IBOutlet weak var comboBoxMonthes: NSComboBox!
    @IBOutlet var LogProcessing: NSTextView!
    
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
    
    func clearLog() {
        self.LogProcessing.string = "";
    }
    
    func addLog(_ line: String) {
        let _line: String = line + "\n"
        let attrLine = NSMutableAttributedString(string: _line)
        let range = (_line as NSString).range(of: _line)
        let color: NSColor
        if let clr = self.defaultTextColor {
            color = clr
        } else {
            color = NSColor.textColor
        }
        attrLine.addAttribute(NSAttributedString.Key.foregroundColor, value: color, range: range)
        self.LogProcessing.textStorage?.append(attrLine)
        self.LogProcessing.scrollToEndOfDocument(nil)
    }
    
    func addLogError(_ line: String) {
        let _line: String = "ОШИБКА: " + line + "\n"
        let attrLine = NSMutableAttributedString(string: _line)
        let range = (_line as NSString).range(of: _line)
        attrLine.addAttribute(NSAttributedString.Key.foregroundColor, value: NSColor.red, range: range)
        self.LogProcessing.textStorage?.append(attrLine)
        self.LogProcessing.scrollToEndOfDocument(nil)
    }
    
    func loadMonthesFromFile() {
        self.months.removeAll()
        self.comboBoxMonthes.removeAllItems()
        self.classifications = nil
        self.clearLog();
        let filePath = self.pathToFile.stringValue
        addLog("Пробую открыть файл \(filePath)")
        guard let file = XLSXFile(filepath: filePath) else {
            addLogError("XSLX сломан или его не существует")
            return
        }
        
        if let workbooks: [Workbook] = try? file.parseWorkbooks() {
            if workbooks.count == 0 {
                addLogError("не найдено ни одного workbook");
                return
            }
            if workbooks.count > 1 {
                addLogError("найдено больше одного workbook");
                return
            }
            let workbook: Workbook = workbooks[0]
            for (worksheetName, path) in try! file.parseWorksheetPathsAndNames(workbook: workbook) {
                if let worksheetName = worksheetName {
                    if PageMonth.isPageMonth(name: worksheetName) {
                        addLog("Найдена страница \(worksheetName)")
                        print("This page is month")
                        months.append(PageMonth(name: worksheetName, path: path))
                    }
                    
                    if PageClassifications.isPageClassifications(name: worksheetName) {
                        addLog("Найдена страница \(worksheetName)")
                        self.classifications = PageClassifications(path: path);
                    }
                }
            }
        }
        if months.count == 0 {
            addLogError("не найдено ни одного страницы месяца. должно быть 'месяц 01', 'месяц 02', 'месяц 03' и так далее");
            return
        }
        
        if self.classifications == nil {
            addLogError("страница 'классификации' не найдена - пожалуйста создайте ее");
            return
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
        
        addLog("Загружаю классификации...");
        if let classPath = self.classifications?.path {
            if let worksheet = try? file.parseWorksheet(at: classPath) {
                for row in worksheet.data?.rows ?? [] {
                    for cell in row.cells {
                        if cell.reference.column.value == "A" {
                            print(cell)
                        }
                    }
                }
            }
        }
        addLog("Пока вроде все в порядке");
    }
}

