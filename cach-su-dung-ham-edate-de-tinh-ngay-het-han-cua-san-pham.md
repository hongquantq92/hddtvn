Cách sử dụng hàm EDATE để tính ngày hết hạn của sản phẩm
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-edate-de-tinh-ngay-het-han-cua-san-pham/)Hàm EDATE trả về giá trị ngày trước hoặc sau ngày đã biết đã xác định trước. Khi áp dụng trong thực tế, hàm EDATE thường được dùng để tính toán ngày đáo hạn (thanh toán), ngày đến hạn của sản phẩm. Mời bạn tìm hiểu cách sử dụng hàm EDATE (có ví dụ cụ thể) …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm EDATE trả về giá trị ngày trước hoặc sau ngày đã biết đã xác định trước. Khi áp dụng trong thực tế, hàm EDATE thường được dùng để tính toán ngày đáo hạn (thanh toán), ngày đến hạn của sản phẩm. Mời bạn tìm hiểu cách sử dụng****hàm EDATE (có ví dụ cụ thể) trong bài viết dưới đây.**


### 1. Cấu trúc hàm EDATE


Cú pháp hàm: **=EDATE(start\_date; months)**


*Trong đó:*




* **Start\_date**: Ngày bắt đầu.

* **Months**: Số tháng trước hoặc sau ngày bắt đầu **start\_date**. Giá trị dương cho đối số months tạo ra ngày trong tương lai, giá trị âm tạo ra ngày trong quá khứ.



*Lưu ý:*




* Nếu **start\_date** không phải là ngày hợp lệ, hàm EDATE trả về giá trị lỗi **#VALUE!** .

* Nếu **months** không phải là số nguyên thì nó bị cắt cụt.



### 2. Cách sử dụng hàm EDATE để tính ngày hết hạn của sản phẩm


Ví dụ ta có 4 loại sản phẩn là rượu, bia, pushmax, coca có ngày sản xuất trong hình dưới. Các loại sản phẩm này có thời hạn sử dụng tương ứng là:




* Rượu có thời hạn sử dụng trong 36 tháng tính từ ngày sản xuất.

* Bia có thời hạn sử dụng trong 30 tháng tính từ ngày sản xuất.

* Pushmax có thời hạn sử dung trong 24 tháng tính từ ngày sản xuất.

* Coca có thời hạn sử dụng trong 20 tháng tính từ ngày sản xuất.



![](https://hddtvn.com/wp-content/uploads/2021/01/DmKPMl3.png)


Áp dụng cấu trúc hàm EDATE kết hợp với hàm IF, ta có công thức tính ngày hết hạn của các sản phẩm như sau:


**=IF(B2=”Rượu”;EDATE(E2;36);IF(B2=”Bia”;EDATE(E2;30);IF(B2=”Pushmax”;EDATE(E2;24);EDATE(E2;20))))**


Sao chép công thức cho các sản phẩm còn lại ta sẽ thu được ngày hết hạn của các loại sản phẩm đó như hình dưới.


[![](https://hddtvn.com/wp-content/uploads/2021/01/D9dUVtK.png)](https://hddtvn.com/wp-content/uploads/2021/01/D9dUVtK.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm EDATE để tính ngày hết hạn của sản phẩm. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm EDATE trả về giá trị ngày trước hoặc sau ngày đã biết đã xác…

