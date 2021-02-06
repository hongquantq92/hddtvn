Cách ghép/gộp nhiều file Excel thành một file duy nhất
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-ghep-gop-nhieu-file-excel-thanh-mot-file-duy-nhat/)Bạn thường phải làm tổng hợp, thống kê các báo cáo, danh sách của các bộ phận khác thành một file tổng? Nếu chỉ copy/paste thủ công thì rất mất thời gian và có thể gây sai sót nếu có nhiều dữ liệu. Làm sao để tổng hợp dữ liệu chính xác và nhanh chóng? Bài …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#### **Bạn thường phải làm tổng hợp, thống kê các báo cáo, danh sách của các bộ phận khác thành một file tổng? Nếu chỉ copy/paste thủ công thì rất mất thời gian và có thể gây sai sót nếu có nhiều dữ liệu. Làm sao để tổng hợp dữ liệu chính xác và nhanh chóng? Bài viết này sẽ hướng dẫn các bạn cách ghép/gộp nhiều file Excel thành một file duy nhất nhé. Các bạn có thể ứng dụng gộp/ghép file để thống kê thành quả công việc, quyết toán thu, chi của cơ quan mình trong một khoảng thời gian bất kỳ.**


![Cách ghép/gộp nhiều file Excel thành một file duy nhất](https://hddtvn.com/wp-content/uploads/2021/01/ghep-file-excel.jpg)


Ví dụ ta có 4 file Excel trong thư mục **Gop file** như hình dưới. Yêu cầu cần gộp cả 4 file Excel này lại thành một file duy nhất. Để thực hiện thì các bạn hãy làm theo các bước sau đây nhé.


![Cách ghép/gộp nhiều file Excel thành một file duy nhất](https://hddtvn.com/wp-content/uploads/2021/01/18-3.png "Cách ghép/gộp nhiều file Excel thành một file duy nhất")


**Bước 1:** Đầu tiên, các bạn cần mở một file Excel mới lên. Sau đó các bạn chọn thẻ **Developer** trên thanh công cụ. Tiếp theo các bạn chọn **Visual Basic** tại mục **Code**. Hoặc các bạn cũng có thể sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ VBA.


![Cách ghép/gộp nhiều file Excel thành một file duy nhất](https://hddtvn.com/wp-content/uploads/2021/01/13-3.png "Cách ghép/gộp nhiều file Excel thành một file duy nhất")


**Bước 2:** Lúc này, cửa sổ Microsoft Visual Basic for Applications hiện ra. Các bạn chọn thẻ **Insert** trên thanh công cụ. Thanh cuộn hiện ra thì các bạn chọn mục **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/14-2.png)


**Bước 3:** Lúc này, hộp thoại Module hiện ra. Các bạn sao chép đoạn code dưới đây vào hộp thoại Module.


Sub copy()  

Path = "C:\Users\admin\OneDrive\Desktop\Gop file\"  

Filename = Dir(Path & "*.xls*")  

Do While Filename <> ""  

Workbooks.Open Filename:=Path & Filename, ReadOnly:=True  

For Each Sheet In ActiveWorkbook.Sheets  

Sheet.copy after:=ThisWorkbook.Sheets(1)  

Next  

Workbooks(Filename).Close  

Filename = Dir()  

Loop  

End Sub


Sau đó các bạn nhấn **Run** trên thanh công cụ hoặc nhấn phím **F5** để chạy mã code.


![](https://hddtvn.com/wp-content/uploads/2021/01/15-2.png)


Lưu ý là đoạn Path = “” : Bên trong dấu ngoặc là đường dẫn của thư mục chứa các file bạn lưu. Và nhớ thêm dấu gạch “\” sau cùng đường dẫn để nó hiểu là 1 folder nhé.


![](https://hddtvn.com/wp-content/uploads/2021/01/19-3.png)


Chỉ cần như vậy là tất cả file Excel trong thư mục **Gop file** đã được gộp lại thành một file Excel duy nhất. Bây giờ thì các bạn cần tiến hành chỉnh sửa dữ liệu trong file Excel mới cũng như nhấn **Save** để lưu lại file này.


![](https://hddtvn.com/wp-content/uploads/2021/01/17-3.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng code VBA để gộp nhiều file Excel thành một file duy nhất. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### 


moreBạn thường phải làm tổng hợp, thống kê các báo cáo, danh sách của các…

