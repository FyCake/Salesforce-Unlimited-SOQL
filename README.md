![image](https://github.com/user-attachments/assets/ef7e928f-abb3-4d9d-80e2-caf4d47fccd8)# Salesforce-Unlimited-SOQL-
当Keys List过大时，可以使用本工具一次性导出。


Salesforce Unlimited SOQL Guide
1.	首先需要安装Salesforce CLI，安装地址：https://developer.salesforce.com/tools/salesforcecli
如果不会配Path，就直接默认安装即可。
![image](https://github.com/user-attachments/assets/b38b0c97-6734-4e45-ba54-cbdf0f65da87)
2.	解压UnlimitedSOQL_QT工具压缩包，得到下图所示文件。（注意：其他程序文件不要删。）
![image](https://github.com/user-attachments/assets/3cdfb67b-c41a-459d-96f3-31f725d68010)
3.	点击dist → main文件夹进去，看到main.exe文件。这就是我们的Unlimited SOQL的小工具。
 ![image](https://github.com/user-attachments/assets/06c55625-cc51-406e-b81e-df82eac20de5)
4.	双击main.exe。
Org Alias Name：你和Salesforce Org（不论哪个Salesforce）连接时你取得名字，随便取，但是要记住，因为下一次再使用工具，如果不想再连一次系统，就记住你取的名字。
Excel文件路径：你想查询的Keys列表。Excel中必须要有Sheet1这张子表，因为工具程序读的就是Sheet1（区分大小写）。Sheet1里面第一行为标题栏，标题栏不要填除标题外的数据，因为不读第一行。
SOQL 查询语句：SELECT FIELD1, FIELD2, FIELD3 FROM object_name WHERE FIELD的形式（不区分大小写）;Where后面最后一个Field就是你Excel Sheet1第一列数据的字段名，其他筛选条件可以放在这个字段前面，这个字段必须放在最后一位。（增加了type为Address的字段导出。）
例如：
SOQL查询语句为：SELECT Id,name,BillingAddress, Active__c,Owner.name,OwnerId,owner.Address,Owner.Alias,CreatedById,CreatedBy.UserRole.name,CreatedDate,CreatedBy.Profile.name,Owner.ProfileId FROM Account where id
 ![image](https://github.com/user-attachments/assets/4f73dad3-0e7d-4176-877f-d1f868e7d310)
5.	点击“执行查询”，如果是电脑之前都没连接过的Alias，那么就会跳转网页让你验证登录。如果连接过的，就不会跳转网页再次验证登录。
然后开始执行查询。
![image](https://github.com/user-attachments/assets/9a3ed263-69f1-4a06-b0e3-be3fb757d1a9)
![image](https://github.com/user-attachments/assets/3ddb989f-ebb6-4d7b-ac5e-37c23b96b257)
6.	日志输出那块可以看运行进度，下图为查询完成。
![image](https://github.com/user-attachments/assets/082a9d49-218d-47e9-a32a-a94b2a3fbacd)
7.	Excel中最后获得三个子表。
Sheet1：你的Keys List；
Sheet1去重后List；
SOQL Result：根据SOQL导出的数据List。（如果表中已经存在子表SOQL Result，程序会新建‘SOQL Result+数字’新的子表存数据，比如下图，对应的SOQL Result2则是我这个Guide演示所Export的数据。）
![image](https://github.com/user-attachments/assets/e36e405d-7499-465d-a3f0-1453543ecec5)

