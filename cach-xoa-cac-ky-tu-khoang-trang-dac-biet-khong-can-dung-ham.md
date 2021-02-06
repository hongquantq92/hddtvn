Cách xóa các ký tự khoảng trắng đặc biệt không cần dùng hàm
===========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-xoa-cac-ky-tu-khoang-trang-dac-biet-khong-can-dung-ham/)Trong Excel có những ký tự khoảng trắng đặc biệt gây lỗi hàm mà dùng mắt thường thì bạn sẽ khó có thể phát hiện được. Ví dụ khi 2 ô chứa đoạn văn bản giống nhau, 1 ô có khoảng trắng còn ô kia không có, thì khoảng trắng dù nhỏ đến đâu, dù …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong Excel có những ký tự khoảng trắng đặc biệt gây lỗi hàm mà dùng mắt thường thì bạn sẽ khó có thể phát hiện được. Ví dụ khi 2 ô chứa đoạn văn bản giống nhau, 1 ô có khoảng trắng còn ô kia không có, thì khoảng trắng dù nhỏ đến đâu, dù chỉ là 1 ký tự cũng khiến 2 ô có giá trị khác nhau. Khi đó, bạn sẽ phải đau đầu tìm ra lý do tại sao công thức đúng lại không hiệu quả với 2 dữ liệu nhìn giống nhau. Cách giải quyết lúc này là xóa các khoảng trắng thừa trong dữ liệu. Người dùng cẩn thận có thể tìm ra khoảng trắng ẩn ở trước đoạn văn bản hoặc vài khoảng trắng thừa ở giữa các từ. Nhưng mắt thường lại gần như không thể tìm ra khoảng trắng phía sau, hoặc khoảng trắng ở cuối ô. Hãy đọc bài viết này để biết cách xóa các ký tự khoảng trắng đặc biệt không cần dùng hàm trong Excel nhé.**


![Cách xóa các ký tự khoảng trắng đặc biệt không cần dùng hàm](https://hddtvn.com/wp-content/uploads/2021/01/xoa-ky-tu-khoang-trang.jpg)


Ví dụ như trong bảng của hình dưới. Ta có thể thấy công thức trong ô F4 là =D4*E4. Trong khi đó ta thấy ô D4 không có gì thì đáng nhẽ ra kết quả trả về tại ô F4 sẽ là 0. Tuy nhiên kết quả lại là lỗi #VALUE!. Điều này có nghĩa là ô D4 có ký tự đặc biệt khoảng trắng. Để giải quyết vấn đề này thì các bạn hãy làm theo các bước dưới đây nhé.


![](https://hddtvn.com/wp-content/uploads/2021/01/2-3.png)


**Bước 1:** Đầu tiên, các bạn cần bôi đen toàn bộ bảng dữ liệu. Sau đó các bạn chọn thẻ **Home** trên thanh công cụ rồi chọn mục **Format as Table**. Thanh cuộn hiện ra thì các bạn chọn một mẫu bảng để chuyển bảng dữ liệu của mình thành chế độ Table.


![](https://hddtvn.com/wp-content/uploads/2021/01/3-3.png)


**Bước 2:** Sau khi bảng dữ liệu đã được chuyển sang chế độ Table. Các bạn chọn thẻ **Data** trên thanh công cụ. Sau đó nhấn vào **From Table** tại mục **Get & Transform**.


![](https://hddtvn.com/wp-content/uploads/2021/01/4-3.png)


#### **Bước 3:** Lúc này, bảng dữ liệu của các bạn đã được chuyển vào cửa sổ Power Query Editor.


Tiếp theo các bạn bôi đen toàn bộ cột có chứa ký tự đặc biệt khoảng trắng. Sau đó các bạn chọn thẻ **Transform** trên thanh công cụ rồi chọn mục **Format**. Thanh cuộn hiện ra thì các bạn chọn mục **Trim**.


![](https://hddtvn.com/wp-content/uploads/2021/01/5-4.png)


**Bước 4:** Tiếp theo, các bạn chọn thẻ Home rồi chọn mục **Data Type**. Thanh cuộn hiện ra thì các bạn chọn mục **Decimal Number** để chuyển dữ liệu về dạng số.


![](https://hddtvn.com/wp-content/uploads/2021/01/6-3.png)


**Bước 5:** Cuối cùng, các bạn nhấn vào dấu **X** ở trên cùng góc bên phải để thoát khỏi cửa sổ Power Query Editor. Lúc này sẽ có thông báo hỏi xem bạn có muốn giữ thay đổi không thì các bạn chọn **Keep** nhé.


![](https://hddtvn.com/wp-content/uploads/2021/01/7-3.png)


Chỉ cần như vậy là các ký tự khoảng trắng trong bảng đã được xử lý và lỗi #VALUE! đã không còn nữa.


![](https://hddtvn.com/wp-content/uploads/2021/01/8-3.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách xóa các ký tự khoảng trắng trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### <a href=”http://


moreTrong Excel có những ký tự khoảng trắng đặc biệt gây lỗi hàm mà dùng…

