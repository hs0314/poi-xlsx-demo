package me.heesu.poixlsx;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class ExcelMngController {

    @Autowired
    ExcelDownloadService excelDownloadService;

    @GetMapping("/download")
    @ResponseBody
    public String excelDownload(){
        excelDownloadService.processExcelDownload();

        return "";
    }

}
