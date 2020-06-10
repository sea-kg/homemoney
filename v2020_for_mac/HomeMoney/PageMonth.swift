//
//  PageMonth.swift
//  HomeMoney
//
//  Created by sea-kg on 11.06.2020.
//  Copyright © 2020 sea-kg. All rights reserved.
//

import Foundation

class PageMonth {
    let name: String
    let path: String
    
    init(name: String, path: String) {
        self.name = name;
        self.path = path;
    }
    
    static func isPageMonth(name: String) -> Bool {
        let range = NSRange(location: 0, length: name.utf16.count)
        let regex = try! NSRegularExpression(pattern: "месяц [0-9]{2}")
        return regex.firstMatch(in: name, options: [], range: range) != nil
    }
}
