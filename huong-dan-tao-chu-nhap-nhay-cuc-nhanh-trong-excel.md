Hướng dẫn tạo chữ nhấp nháy (cực nhanh) trong Excel
===================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-tao-chu-nhap-nhay-cuc-nhanh-trong-excel/)Khi sử dụng internet thường xuyên chắc hẳn bạn sẽ gặp các chữ nhấp nháy trên các trang web hoặc blog… Mục đích của việc tạo chữ nhấp nháy là để làm nổi bật hơn, qua đó gây chú ý cho độc giả và người đọc đến vị trí mà bạn muốn hướng tới. Tương …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi sử dụng internet thường xuyên chắc hẳn bạn sẽ gặp các chữ nhấp nháy trên các trang web hoặc blog… Mục đích của việc tạo chữ nhấp nháy là để làm nổi bật hơn, qua đó gây chú ý cho độc giả và người đọc đến vị trí mà bạn muốn hướng tới. Tương tự như vậy, bạn có thể tạo chữ nhấp nháy nếu muốn nhấn mạnh hoặc làm bắt mắt hơn cho dữ liệu trong bảng tính Excel. Trong bài viết này, hddtvn sẽ hướng dẫn các bạn cách làm nhé.**


![Hướng dẫn tạo chữ nhấp nháy (cực nhanh) trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/nhap-nhay.jpg)


**Bước 1:** Để tạo chữ nhấp nháy, đầu tiên các bạn cần mở file Excel cần tạo chữ nhấp nháy lên. Sau đó các bạn chọn thẻ **Developer** trên thanh công cụ. Tiếp theo các bạn chọn **Visual Basic** tại mục **Code**. Hoặc các bạn cũng có thể sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ VBA.


![](https://hddtvn.com/wp-content/uploads/2021/01/MwhlBVE.png)


**Bước 2:** Lúc này, cửa sổ VBA hiện ra. Các bạn chọn thẻ **Insert** => **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/2oZhIDC.png)


**Bước 3:** Tiếp theo, các bạn sao chép đoạn code dưới đây vào hộp thoại Module.


Sub StartBlink()  

Dim xCell As Range  

Dim xTime As Variant  

Set xCell = Range("A1")  

With ThisWorkbook.Worksheets("Sheet1").Range("A1").Font  

If xCell.Font.Color = vbRed Then  

xCell.Font.Color = vbWhite  

Else  

xCell.Font.Color = vbRed  

End If  

End With  

xTime = Now + TimeSerial(0, 0, 1)  

Application.OnTime xTime, "'" & ThisWorkbook.Name & "'!StartBlink", , True  

End Sub


Trong đó **A1** là vị trí ô bạn muốn tạo chữ nhấp nháy, **Sheet1** là trang tính mà bạn muốn tạo chữ. Các bạn có thể thay đổi đối số này tùy thuộc vào vị trí mà bạn muốn tạo chữ nhấp nháy.


![](https://hddtvn.com/wp-content/uploads/2021/01/MSjxEW4.png)


**Bước 4:** Bây giờ quay lại trang tính, các bạn nhập nội dung vào ô **A1**. Sau đó chọn thẻ **Developer** trên thanh công cụ => **Insert** => biểu tượng của **Button**.


![](https://hddtvn.com/wp-content/uploads/2021/01/rCDFteT.png)


#### **Bước 5:** Lúc này, hộp thoại Assign Macro hiện ra.


Tại mục **Macro name** các bạn chọn **StartBlink** và tại mục **Macros in** các bạn chọn **This Workbook**.


![](https://hddtvn.com/wp-content/uploads/2021/01/3gT0fmb.png)


**Bước 6:** Quay lại trang tính, các bạn nhấn và kéo chuột trái để tại nút bấm rồi nhấn chuột phải vào nút rồi và chọn **Edit Text** để chỉnh sửa nội dung bên trong.


![](https://hddtvn.com/wp-content/uploads/2021/01/UmsIj6k.png)


**Bước 7:** Chỉ cần như vậy là ta đã tạo xong nút nhấn Bắt đầu/Tạm dừng nhấp nháy. Chỉ cần nhấn vào nút này là chữ trong ô A1 sẽ nhấp nháy hoặc tạm dừng.


![](https://hddtvn.com/wp-content/uploads/2021/01/tecY7Xu.png)


**Bước 8:** Các bạn còn có thể chỉnh màu nhấp nháy bằng cách nhấn **Alt + F11**. Sau đó các bạn có thể chỉnh màu nhấp nháy tại mục **vbRed** và **vbGreen** trong đoạn code. Các bạn chỉ cần nhập màu mà mình muốn vào đây là được.


![](https://hddtvn.com/wp-content/uploads/2021/01/TyRuKvw.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tạo chữ nhấp nháy trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreKhi sử dụng internet thường xuyên chắc hẳn bạn sẽ gặp các chữ nhấp nháy…

