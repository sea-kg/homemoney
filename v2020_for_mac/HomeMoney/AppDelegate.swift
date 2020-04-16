//
//  AppDelegate.swift
//  HomeMoney
//
//  Created by sea-kg on 01.04.2020.
//  Copyright © 2020 sea-kg. All rights reserved.
//

import Cocoa

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
            }
        }
    }
}

