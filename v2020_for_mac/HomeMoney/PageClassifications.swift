//
//  PageClassifications.swift
//  HomeMoney
//
//  Created by sea-kg on 11.06.2020.
//  Copyright © 2020 sea-kg. All rights reserved.
//

import Foundation

class PageClassifications {
    let path: String
    
    init(path: String) {
        self.path = path;
    }
    
    static func isPageClassifications(name: String) -> Bool {
        return name == "классификации"
    }
}
