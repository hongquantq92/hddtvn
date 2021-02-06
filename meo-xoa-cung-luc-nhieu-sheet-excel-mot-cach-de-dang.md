Mẹo xóa cùng lúc nhiều Sheet Excel một cách dễ dàng
===================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/meo-xoa-cung-luc-nhieu-sheet-excel-mot-cach-de-dang/)Xóa Sheet là điều thường gặp khi làm việc với bảng tính Excel. Nếu cần xóa một vài sheet thì ta có thể làm thao tác thủ công. Nhưng cần phải xóa hàng chục, thậm chí hàng trăm Sheet (để bảng tính đỡ rối) thì sẽ là cả một vấn đề lớn nếu làm theo …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Xóa Sheet là điều thường gặp khi làm việc với bảng tính Excel. Nếu cần xóa một vài sheet thì ta có thể làm thao tác thủ công. Nhưng cần phải xóa hàng chục, thậm chí hàng trăm Sheet (để bảng tính đỡ rối) thì sẽ là cả một vấn đề lớn nếu làm theo cách thông thường. Mời các bạn theo dõi bài viết sau để biết cách xóa cùng lúc nhiều Sheet Excel nhé.**


### 1. Xóa cùng lúc nhiều sheet Excel bằng cách Group Sheet


Thông thường, để xóa một sheet Excel thì chúng ta chỉ cần nhấn chuột phải vào sheet đó rồi chọn **Delete** là được.


![](https://hddtvn.com/wp-content/uploads/2021/01/BKhykti.png)


Còn nếu bạn muốn xóa cùng lúc nhiều sheet thì các bạn chỉ cần nhấn và giữ phím **Ctrl** rồi nhấn chọn tất cả những sheet mình muốn xóa để nhóm chúng vào một Group. Sau đó các bạn nhấn chuột phải vào một sheet bất kỳ rồi chọn **Delete**. Chỉ cần như vậy là tất cả sheet được chọn sẽ được xóa đi một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/YwLmTHH.png)


Hoặc nếu bạn muốn xóa tất cả sheet có trong file thì các bạn chỉ cần nhóm tất cả sheet có trong file vào một Group bằng cách nhấn chuột phải vào một sheet bất kỳ rồi chọn **Select All Sheets**. Sau đó các bạn nhấn chuột phải và chọn **Delete** như các cách trên là tất cả sheet trong file sẽ được xóa đi một cách dễ dàng.


![](https://hddtvn.com/wp-content/uploads/2021/01/JLkvQ18.png)


### 2. Xóa cùng lúc nhiều sheet bằng code VBA


Với phương pháp này các bạn có thể xóa hết tất cả các sheet có trong file Excel của mình chỉ trừ những sheet được đặt tên là **GPE**. Vì vậy các bạn phải đặt tên **GPE** cho những sheet muốn giữ lại trong file Excel của mình. Sau đó, các bạn chọn thẻ **Developer** => **Visual Basic** hoặc nhấn tổ hợp phím tắt **Alt + F11** để bật **Microsoft Visual Basic for Applications** lên.


![](https://hddtvn.com/wp-content/uploads/2021/01/JxByUXI.png)


Tại cửa sổ VBA, các bạn chọn thẻ **Insert** => **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/po6wCvB.png)


Sau đó, các bạn sao chép đoạn code sau vào hộp thoại Module:


 Public Sub GPE()  

Application.DisplayAlerts = False  

Dim Ws As Worksheet  

For Each Ws In Worksheets  

If Ws.Name <> "GPE" Then Ws.Delete  

Next Ws  

Application.DisplayAlerts = True  

End Sub


Sau đó các bạn nhấn **Run** hoặc nhấn **F5** để chạy mã code. Chỉ cần như vậy là những sheet Excel không có tên là GPE sẽ được xóa đi một cách nhanh chóng.


[![](https://hddtvn.com/wp-content/uploads/2021/01/IUjh1vu.png)](https://hddtvn.com/wp-content/uploads/2021/01/IUjh1vu.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách xóa cùng lúc nhiều sheet Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreXóa Sheet là điều thường gặp khi làm việc với bảng tính Excel. Nếu cần…

