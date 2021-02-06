Hướng dẫn tách các Sheet thành từng file Excel riêng biệt
=========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-tach-cac-sheet-thanh-tung-file-excel-rieng-biet/)Trên một file Excel bạn có thể tạo nhiều sheet để đễ theo dõi và chuyển đổi giữa các nội dung liên quan tới nhau. Nhưng nhiều sheet cũng làm tăng dung lượng và khiến mở file lâu hơn, nhất là khi bạn sử dụng thêm các add ins hay các code macro VBA. Giải …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#### **Trên một file Excel bạn có thể tạo nhiều sheet để đễ theo dõi và chuyển đổi giữa các nội dung liên quan tới nhau. Nhưng nhiều sheet cũng làm tăng dung lượng và khiến mở file lâu hơn, nhất là khi bạn sử dụng thêm các add ins hay các code macro VBA. Giải pháp lúc này là bạn nên tách các sheet thành từng file Excel riêng biệt. Nhờ vậy nội dung trong sheet vẫn được giữ nguyên, tên sheet trở thành tên của từng file và công thức sử dụng trong sheet không bị thay đổi hay lỗi kết quả. Bài viết dưới đây sẽ hướng dẫn bạn đọc cách tách từng sheet trong Excel thành các file riêng biệt nhé.**


![Hướng dẫn tách các Sheet thành từng file Excel riêng biệt](https://hddtvn.com/wp-content/uploads/2021/01/tach-sheet-tung-file-copy.jpg "Hướng dẫn tách các Sheet thành từng file Excel riêng biệt")


Ví dụ ta có file Excel chứa các sheet như hình dưới. Yêu cầu cần tách các sheet đó thành từng file Excel riêng biệt. Các bạn hãy làm theo các bước dưới đây để tách sheet nhé.


![Hướng dẫn tách các Sheet thành từng file Excel riêng biệt](https://hddtvn.com/wp-content/uploads/2021/01/28.png "Hướng dẫn tách các Sheet thành từng file Excel riêng biệt")


**Bước 1:** Để tách sheet thành từng file Excel riêng biệt, đầu tiên thì các bạn cần chọn thẻ Developer trên thanh công cụ. Sau đó các bạn chọn **Visual Basic** tại mục **Code**. Hoặc các bạn có thể sử dụng tổ hợp phím tắt **Alt + F11** để mở nhanh cửa sổ VBA.


![](https://hddtvn.com/wp-content/uploads/2021/01/29.png)


**Bước 2:** Lúc này, cửa sổ Microsoft Visual Basic for Applications hiện ra. Các bạn chọn thẻ **Insert** trên thanh công cụ. Thanh cuộn hiện ra thì các bạn chọn mục **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/30-1.png)


**Bước 3:** Lúc này, hộp thoại Module hiện ra. Các bạn sao chép mã code dưới đây vào hộp thoại đó.


Sub Splitbook()  

'Updateby20140612  

Dim xPath As String  

xPath = Application.ActiveWorkbook.Path  

Application.ScreenUpdating = False  

Application.DisplayAlerts = False  

For Each xWs In ThisWorkbook.Sheets  

xWs.Copy  

Application.ActiveWorkbook.SaveAs Filename:=xPath & "\" & xWs.Name & ".xls"  

Application.ActiveWorkbook.Close False  

Next  

Application.DisplayAlerts = True  

Application.ScreenUpdating = True  

End Sub


Sau đó các bạn nhấn vào biểu tượng của **Run** ở trên thanh công cụ để chạy mã code này.


![](https://hddtvn.com/wp-content/uploads/2021/01/32.png)


**Bước 4:** Sau khi chạy mã code, kết quả ta sẽ thu được là mỗi sheet đã được tách riêng ra thành một file Excel riêng biệt mà vẫn giữ nguyên dữ liệu.


![](https://hddtvn.com/wp-content/uploads/2021/01/33.png)


Khi các bạn mở những file Excel tách riêng từng sheet đó ra thì sẽ có thông báo như hình dưới hiện lên. Các bạn chỉ cần nhấn **Yes** là có thể chỉnh sửa dữ liệu một cách bình thường.


![](https://hddtvn.com/wp-content/uploads/2021/01/34.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tách sheet thành từng file Excel riêng biệt. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrên một file Excel bạn có thể tạo nhiều sheet để đễ theo dõi và…

