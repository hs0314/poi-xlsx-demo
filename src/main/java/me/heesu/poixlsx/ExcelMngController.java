package me.heesu.poixlsx;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;

@Controller
public class ExcelMngController {

    @Autowired
    ExcelCreateService excelCreateService;

    @Autowired
    ExcelDownloadService excelDownloadService;

    @PostMapping("/create")
    @ResponseBody
    public String excelCreate(@RequestBody HashMap<String, Integer> vo){
        excelCreateService.createXlsxExcel(vo.get("reqRowSize").intValue());

        return "Done!";
    }

    @GetMapping("/download")
    @ResponseBody
    public String excelDownload(HttpServletResponse response){
        try {
            excelDownloadService.createXlsxExcelFileBySax(response);
        }catch(Exception e){

        }

        return "";
    }

}
