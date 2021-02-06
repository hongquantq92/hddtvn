Cách đặt mật khẩu bảo vệ cho nhiều sheet Excel cùng lúc
=======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dat-mat-khau-bao-ve-cho-nhieu-sheet-excel-cung-luc-2/)Đặt mật khẩu bảo vệ sheet Excel sẽ vô cùng cần thiết nếu bạn không muốn cho người khác chỉnh sửa dữ liệu của mình. Tuy nhiên nếu file Excel có nhiều sheet thì việc đặt mật khẩu cho toàn bộ sheet sẽ tốn rất nhiều thời gian và thao tác nếu làm thủ công. …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Đặt mật khẩu bảo vệ sheet Excel sẽ vô cùng cần thiết nếu bạn không muốn cho người khác chỉnh sửa dữ liệu của mình. Tuy nhiên nếu file Excel có nhiều sheet thì việc đặt mật khẩu cho toàn bộ sheet sẽ tốn rất nhiều thời gian và thao tác nếu làm thủ công. Mời bạn theo dõi bài viết sau để biết cách đặt mật khẩu bảo vệ cho nhiều sheet Excel cùng lúc nhé.**


![Cách đặt mật khẩu bảo vệ cho nhiều sheet Excel cùng lúc](https://hddtvn.com/wp-content/uploads/2021/01/mat-khau-sheet.png)


### 1. Đặt mật khẩu bảo vệ sheet Excel theo cách thông thường


Thông thường, nếu muốn đặt mật khẩu bảo vệ sheet Excel thì các bạn cần mở sheet đó lên rồi chọn thẻ **Review** trên thanh công cụ. Sau đó các bạn chọn mục **Protect Sheet** tại mục **Protect**.


![](https://hddtvn.com/wp-content/uploads/2021/01/33-3.png)


Hoặc các bạn có thể nhấn chuột phải vào sheet đó rồi chọn **Protect Sheet**.


![](https://hddtvn.com/wp-content/uploads/2021/01/34-2.png)


Lúc này, hộp thoại Protect Sheet hiện ra. Các bạn nhập mật khẩu bảo vệ vào mục **Password to unprotect sheet** rồi nhấn **OK**. Như vậy là sheet hiện tại sẽ được tạo mật khẩu bảo vệ. Các bạn tiếp tục lặp lại các bước này đối với tất cả các sheet còn lại mà bạn muốn tạo mật khẩu. Tuy nhiên việc này sẽ rất tốn thời gian nếu bạn muốn tạo mật khẩu cho nhiều sheet cùng lúc.


![](https://hddtvn.com/wp-content/uploads/2021/01/35-3.png)


### 2. Đặt mật khẩu bảo vệ cho nhiều sheet Excel cùng lúc bằng VBA


Để đặt mật khẩu bảo vệ cho nhiều sheet Excel cùng lúc bằng VBA, đầu tiên các bạn cần chọn thẻ **Developer** trên thanh công cụ. Sau đó các bạn nhấn chọn **Visual Basic** tại mục **Code**. Hoặc các bạn có thể sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ VBA.


![](https://hddtvn.com/wp-content/uploads/2021/01/36-2.png)


Lúc này, cửa sổ Microsoft Visual Basic for Applications hiện ra. Các bạn chọn thẻ **Insert** trên thanh công cụ. Thanh cuộn hiện ra thì các bạn chọn mục **Module.**


![](https://hddtvn.com/wp-content/uploads/2021/01/MM1ttY8.png)


Tiếp theo, các bạn sao chép đoạn code dưới đây vào cửa sổ Module.


Sub Protect\_Unprotect\_Ws()  

Dim Ws As Worksheet  

For Each Ws In Worksheets  

Ws.Protect Password:="ketoan.vn"  

Next Ws  

Set Ws = Nothing  

End Sub


Trong đó các bạn nhập mật khẩu muốn tại vào mục **Password:=”ketoan.vn”** . Sau đó các bạn nhấn vào mục **Run** trên thanh công cụ hoặc nhấn phím **F5** để chạy mã code. Chỉ cần như vậy là tất cả sheet trong file Excel của các bạn đã được tạo mật khẩu bảo vệ.


![](https://hddtvn.com/wp-content/uploads/2021/01/38-2.png)


Nếu bạn không muốn đặt mật khẩu nữa thì có thể mở khóa tất cả các sheet bằng cách nhập đoạn mã sau vào hộp thoại Module.


Sub Protect\_Unprotect\_Ws()  

Dim Ws As Worksheet  

For Each Ws In Worksheets  

Ws.Unprotect Password:="ketoan.vn"  

Next Ws  

Set Ws = Nothing  

End Sub


Sau đó các bạn nhấn vào mục **Run** trên thanh công cụ hoặc nhấn phím **F5** để chạy mã code. Chỉ cần như vậy là tất cả sheet trong file Excel của các bạn sẽ được mở khóa.


![](https://hddtvn.com/wp-content/uploads/2021/01/39-1-1.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách đặt mật khẩu bảo vệ cho nhiều sheet Excel cùng lúc. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### 


moreĐặt mật khẩu bảo vệ sheet Excel sẽ vô cùng cần thiết nếu bạn không…

