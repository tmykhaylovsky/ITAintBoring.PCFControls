﻿/// <reference path="../lib/@types/xrm/index.d.ts" />
namespace Account { 
    export function getClientUrl() {
        var url = Xrm.Page.context.getClientUrl();
        let b = 1;  
        alert(url);
    }
}  