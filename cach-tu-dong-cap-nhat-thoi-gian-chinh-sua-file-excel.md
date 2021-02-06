Cách tự động cập nhật thời gian chỉnh sửa file Excel
====================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-tu-dong-cap-nhat-thoi-gian-chinh-sua-file-excel/)Nếu cả nhóm làm việc chung trên một file Excel thì nội dung file sẽ được chỉnh sửa nhiều lần, trong khoảng thời gian khác nhau. Để dễ quản lý nội dung file Excel hoặc nắm rõ thời điểm người khác đã chỉnh sửa file, bạn nên chèn thêm thời gian tự động cập nhật …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Nếu cả nhóm làm việc chung trên một file Excel thì nội dung file sẽ được chỉnh sửa nhiều lần, trong khoảng thời gian khác nhau. Để dễ quản lý nội dung file Excel hoặc nắm rõ thời điểm người khác đã chỉnh sửa file, bạn nên chèn thêm thời gian tự động cập nhật vào Excel. Khi đó thông tin về thời gian tạo, chỉnh sửa file gần nhất đều được ghi lại và hiển thị ngay trong Excel. Từ đó giúp bạn nắm rõ về thời điểm cuối cùng nội dung được thay đổi. Mời bạn theo dõi bài viết sau để biết cách tự động cập nhật thời gian chỉnh sửa file Excel nhé.**


**Bước 1:** Đầu tiên, các bạn cần mở Excel cần tự động cập nhật thời gian chỉnh sửa lên. Sau đó các bạn chọn sheet đặt thời gian cập nhật sửa file. Tiếp theo các bạn sử dụng tổ hợp phím tắt **Alt + F11** để mở cửa sổ **Microsoft Visual Basic for Applications**.


**Bước 2:** Sau khi cửa sổ VBA hiện lên. Các bạn chọn thẻ trên thanh công cụ. Thanh cuộn hiện ra các bạn chọn mục **Module**.


![](https://hddtvn.com/wp-content/uploads/2021/01/fK8Gd7H.png)


#### **Bước 3:** Sau khi hộp thoại Module hiện ra, các bạn copy đoạn code sau vào đó:


Sub Workbook\_Open()  

Range("A20").Value = Format(ThisWorkbook.BuiltinDocumentProperties("Creation Date"), "short date")  

Range("B20").Value = Format(ThisWorkbook.BuiltinDocumentProperties("Last Save Time"), "short date")  

End Sub


*Trong đó:*




* A20 là vị trí xuất hiện thời gian tạo file

* B20 là vị trí xuất hiện thời gian chỉnh sửa file gần nhất.



**Bước 4:** Sau khi đã nhập code xong, các bạn nhấn vào biểu tượng của **Run Sub/UserForm** trên thanh công cụ. Hoặc các bạn có thể sử dụng phím tắt **F5** để chạy đoạn code.


![](https://hddtvn.com/wp-content/uploads/2021/01/yxxs9X0.png)


**Bước 5:** Sau khi chạy đoạn code VBA xong. Bây giờ quay lại trang tính Excel, thời gian tạo file và thời gian chỉnh sửa gần nhất đã được xuất hiện tại ô **A20** và **B20**. Bây giờ các bạn có thể chỉnh sửa format cũng như chú thích cho các thời gian đó nhằm dễ hiểu là được.


Chỉ cần như vậy là mỗi lần có người chỉnh sửa file Excel này. Thời gian tại ô B20 sẽ được tự động cập nhật theo đó một cách nhanh chóng. Các bạn có thể từ đó mà quản lý được ai là người chỉnh sửa dữ liệu cũng như thời gian chỉnh sửa.


![](https://hddtvn.com/wp-content/uploads/2021/01/djkWBNQ.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tự động cập nhật thời gian chỉnh sửa file Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreNếu cả nhóm làm việc chung trên một file Excel thì nội dung file sẽ…

