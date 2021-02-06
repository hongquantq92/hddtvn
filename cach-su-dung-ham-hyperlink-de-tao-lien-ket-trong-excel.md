Cách sử dụng hàm HYPERLINK để tạo liên kết trong Excel
======================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-hyperlink-de-tao-lien-ket-trong-excel/)Trong bài viết này, hddtvn sẽ chia sẻ với bạn đọc cách sử dụng hàm HYPERLINK để tạo liên kết trong Excel một cách đơn giản. 1. Cấu trúc hàm HYPERLINK Cú pháp hàm: =HYPERLINK(link\_location; [friendly\_name]) Trong đó:  Link\_location: đối số bắt buộc, là đường dẫn và tên tệp đến tài liệu được mở. Link\_location …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn sẽ chia sẻ với bạn đọc cách sử dụng hàm HYPERLINK để tạo liên kết trong Excel một cách đơn giản.**


### 1. Cấu trúc hàm HYPERLINK


Cú pháp hàm: **=HYPERLINK(link\_location; [friendly\_name])**


*Trong đó:*




* **Link\_location**: đối số bắt buộc, là đường dẫn và tên tệp đến tài liệu được mở. **Link\_location** có thể tham chiếu tới một vị trí trong Excel chẳng hạn như một ô cụ thể hoặc một phạm vi đã đặt tên trong wordbook, hoặc tới một thẻ đánh dấu trong tài liệu Microsoft Word. Đường dẫn có thể đến một tệp được lưu trữ trên ổ đĩa cứng. Đường dẫn cũng có thể là một đường dẫn UNC trên máy chủ hoặc đườn dẫn URL trên Internet hay trên mạng nội bộ.

* **Friendly\_name:** đối số tùy chọn, là nội dung liên kết ( cũng chính là jump text or anchor text) được hiển thị trong ô tính. Nếu bạn bỏ qua thì đối số **link\_location** sẽ được hiển thị là nội dung liên kết luôn.



*Lưu ý:*




* Nếu đường dẫn liên kết không tồn tại hoặc bị hỏng, hàm Hyperlink sẽ báo lỗi khi bạn nhấp chuột vào ô tính.

* Nếu **friendly\_name** trả về giá trị lỗi (ví dụ #VALUE!), thì ô sẽ hiển thị lỗi thay vì hyperlink.

* **Link\_location** có thể là chuỗi văn bản nằm giữa dấu nháy kép hoặc là một tham chiếu đến một ô có chứa nối kết ở dạng chuỗi văn bản.



### 2. Cách sử dụng hàm HYPERLINK


Ví dụ ta cần tạo liên kết tới trang hddtvn thì áp dụng cấu trúc hàm bên trên ta có công thức tạo liên kết như sau:


**=HYPERLINK(“https://www.ketoan.vn/”;”Ketoan.vn”)**


Chỉ cần như vậy là các bạn đã tạo được liên kết tới trang Ketoan.vn. Khi nhấn vào ô vừa tạo, các bạn sẽ được chuyển tới trang hddtvn một cách nhanh chóng.


![](https://hddtvn.com/wp-content/uploads/2021/01/wWNs3Lb.png)


Hoặc nếu ta muốn tạo liên kết tới Sheet2 trong file Excel hiện tại thì các bạn chỉ cần áp dụng công thức sau:


**=HYPERLINK(“#Sheet2!A1″;”Sheet2”)**


Chỉ cần như vậy là khi nhấn vào ô vừa tạo là màn hình sẽ được chuyển sang ô A1 tại Sheet2.


![](https://hddtvn.com/wp-content/uploads/2021/01/AqGpv4M.png)


Hoặc bạn cũng có thể tạo liên kết tới một file khác theo cấu trúc sau:


**=HYPERLINK(“C:\Users\admin\OneDrive\Desktop\check file.xlsx”;”File mới”)**


Chỉ cần như vậy là bạn đã tạo được liên kết tới file Excel khác. Khi nhấn vào ô vừa tạo thì Excel sẽ hiện lên thông báo hỏi xem bạn có muốn chuyển tới đường dẫn đó không. Các bạn chỉ cần nhấn **Yes** là Excel sẽ tự động mở file tại đường dẫn ra.


![](https://hddtvn.com/wp-content/uploads/2021/01/ffsanNf.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm HYPERLINK để tạo liên kết trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong bài viết này, Ketoan.vn sẽ chia sẻ với bạn đọc cách sử dụng hàm…

