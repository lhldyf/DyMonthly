import com.lhldyf.common.utils.ExcelUtils;
import com.lhldyf.common.utils.FileUtils;

import java.time.LocalTime;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main {

    private static List<String> readFromNewTemplate(String fileName) {
        List<List<String>> list =  ExcelUtils.readSheet(fileName,1);
        if(list.size() == 3) {
            return list.get(2);
        } else if(list.size() == 21) {
            return list.get(3);
        } else {
            System.out.println("文件读取出来的值不是3行也不是21行，请检查文件内容是否正常。文件名+"+fileName+",总行数："+list.size());
            return null;
        }
    }

    public static void main(String[] args) {
        /*List<List<String>> list = new ArrayList<>();
        List<String> oneRow = readFromNewTemplate("/Users/LH/Desktop/2017-06-27/北京丰宏智达人力资源有限公司：信息登记表.docx.xlsx");
        if(null != oneRow) {
            list.add(oneRow);
        }
        ExcelUtils.writeToExcel(list,"/Users/LH/Desktop/2017-06-27/test.xlsx");
        System.out.println("执行完成");*/

        List<String> filePathList = FileUtils.getFilePathList("/Users/LH/Desktop/2017-06-27");
        Map<String,List<String>> map = new HashMap<>();

        for(String filePath : filePathList) {
            List<String> oneRow = readFromNewTemplate(filePath);
            if(null != oneRow) {
                map.put(filePath,oneRow);
            }
        }

        System.out.println("共识别"+filePathList.size()+"个文件，其中识别为有效文件的有"+map.size()+"个，下面将这些文件写入新的Excel文件。");

        String exportFileName = LocalTime.now().toString();
        ExcelUtils.writeToExcel(map,"/Users/LH/Desktop/DyMonthLy-"+exportFileName+".xlsx");

        System.out.println("执行完毕！");
    }
}
