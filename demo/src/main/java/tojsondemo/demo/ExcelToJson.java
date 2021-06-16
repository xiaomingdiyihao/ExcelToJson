package tojsondemo.demo;


import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.*;

public class ExcelToJson {
    //json的key
    private static String[] titleArray = {  "type","number","companyName","level","location"};


    public static void main(String[] args) {

        InputStream is = null;
        List<List<String>> excelToJson = excelToJson(is);
        System.out.println("\n数据输出：\n");
        System.out.println((JSON.toJSON(excelToJson)));

    }
    public static List<List<String>> excelToJson(InputStream is) {
        JSONObject jsonRow = new JSONObject();
        List<List<String>> jsonRows = new ArrayList<>();
        try {
            // 获取文件
            is = new FileInputStream("E:\\公司文档\\驾驶舱管理平台\\人才创业项目明细.xlsx");
            // 创建excel工作空间
            Workbook workbook = WorkbookFactory.create(is);
            // 获取sheet页的数量
            int sheetCount = workbook.getNumberOfSheets();

            // 遍历每个sheet页
            for (int i = 0; i < sheetCount; i++) {
                // 获取每个sheet页
                Sheet sheet = workbook.getSheetAt(i);
                // 获取sheet页的总行数
                int rowCount = sheet.getPhysicalNumberOfRows();
                // 遍历每一行
                for (int j = 0; j < rowCount; j++) {
                    //记录每一行的一个唯一标识作为json的key
                    String key="";
                    // 获取列对象
                    Row row = sheet.getRow(j);
                    // 获取总列数
                    int cellCount = row.getPhysicalNumberOfCells();
                    List<String> array = new ArrayList<>();
                    Map<String, String> map = new HashMap<String, String>();
                    //除去第一行的标题
                    if (j > 0) {
                        // 遍历所有数据
                        for (int n = 0; n < cellCount; n++) {
                            // 获取每一个单元格
                            Cell cell = row.getCell(n);
                            // 在读取单元格内容前,设置所有单元格中内容都是字符串类型
                            cell.setCellType(Cell.CELL_TYPE_STRING);
                            // 按照字符串类型读取单元格内数据
                            String cellValue = cell.getStringCellValue();
                            if(n==1) {
                                key=cellValue;
                            }
                            array.add(cellValue);
                            map.put(titleArray[n], cellValue);

                        }
                    }

                    if (map.size() != 0) {
                        jsonRows.add(array);
                        jsonRow.put(key,map);

                    }

                }

            }
        } catch (Exception e) {

        }
        return jsonRows;

    }



    public static void createJsonFile(String jsonData) throws IOException {
        File file = new File("C:\\Users\\Administrator\\Desktop\\用户.json");
        //连接流
        FileOutputStream fos=null;
        //链接流
        BufferedOutputStream bos =null;
        try {
            fos = new FileOutputStream(file);
            bos = new BufferedOutputStream(fos);
            bos.write(jsonData.getBytes());;
            bos.flush();

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }finally {
            if(null!=bos) {
                bos.close();
            }
            if(null!=fos) {
                fos.close();
            }
        }
    }


}
