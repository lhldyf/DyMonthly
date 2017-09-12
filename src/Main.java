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
        } else if(list.size() > 3){
            if(list.get(2).size() != 16 )
            System.out.println("这行数据已经加入统计,请您看下是否正常. fileName:"+fileName + "数据："+list.get(2));
            return list.get(2);
        } else{

            int i = list.size();

            // 表格有的有个隐藏的sheet3页面，所以跳过去读第三页
            list =  ExcelUtils.readSheet(fileName,2);
            if(list.size() == 3) {
                return list.get(2);
            } else if(list.size() == 21) {
                return list.get(3);
            } else if(list.size() > 3){
                if(list.get(2).size() != 16 )
                System.out.println("这行数据已经加入统计,请您看下是否正常. fileName:"+fileName + "数据："+list.get(2));
                return list.get(3);
            }

            System.out.println("文件读取出来的值不是3行也不是21行，请检查文件内容是否正常。文件名:"+fileName+",sheet(1)总行数："+i+"，sheet(2)总行数："+list.size());
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

        List<String> filePathList = FileUtils.getFilePathList("C:\\Users\\Administrator\\Desktop\\dy\\2017-9-12");
        Map<String,List<String>> map = new HashMap<>();

        for(String filePath : filePathList) {
            List<String> oneRow = readFromNewTemplate(filePath);
            if(null != oneRow) {
                map.put(filePath,oneRow);
            }
        }

        System.out.println("共识别"+filePathList.size()+"个文件，其中识别为有效文件的有"+map.size()+"个，下面将这些文件写入新的Excel文件。");

        String exportFileName = LocalTime.now().toString().replace(":","");
        ExcelUtils.writeToExcel(map,"C:\\Users\\Administrator\\Desktop\\dy\\DyMonthLy-"+exportFileName+".xlsx");

        System.out.println("执行完毕！");
    }
}
