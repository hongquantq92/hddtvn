Sử dụng hàm SUMPRODUCT để tính tổng doanh thu/tổng theo điều kiện
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/su-dung-ham-sumproduct-de-tinh-tong-doanh-thu-tong-theo-dieu-kien/)Hàm SUMPRODUCT là sự kết hợp của hàm SUM (tính tổng) và hàm PRODUCT (tính tích). Nó giúp bạn tính tích các phạm vi hoặc mảng với nhau và trả về tổng số sản phẩm. Hàm này được dùng linh hoạt và cực kỳ hữu ích khi bạn cần tính tổng doanh thu, tổng vốn …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm SUMPRODUCT là sự kết hợp của hàm SUM (tính tổng) và hàm PRODUCT (tính tích). Nó giúp bạn tính tích các phạm vi hoặc mảng với nhau và trả về tổng số sản phẩm. Hàm này được dùng linh hoạt và cực kỳ hữu ích khi bạn cần tính tổng doanh thu, tổng vốn đầu tư, tổng giá nguyên vật liệu… hoặc tính tổng của nhiều điều kiện khác. Trong bài viết sau đây, hddtvn sẽ giới thiệu đến bạn đọc cách sử dụng hàm SUMPRODUCT để tính tổng các phép nhân trong Excel.**


### 1. Cấu trúc hàm SUMPRODUCT


Cú pháp hàm: **=SUMPRODUCT(array1; [array2]; [array3]; …)**


*Trong đó:*




* **array1**: đối số bắt buộc, là mảng đầu tiên mà bạn muốn nhân các thành phần của nó rồi cộng tổng.

* **array2, array3,…**: đối số tùy chọn, là các mảng tiếp theo mà bạn muốn nhân các thành phần của nó rồi cộng tổng.



*Lưu ý:*




* Các mảng **array1, array2,…** phải có cùng kích cỡ. Nếu không hàm sẽ trả về lỗi **#VALUE!**.

* Các giá trị trong **array1, array2,…** phải là số. Nếu không hàm sẽ trả về lỗi **#VALUE!**.

* Hàm SUMPRODUCT cho phép nhập tối đa 255 đối số.



### 2. Cách sử dụng hàm SUMPRODUCT


Ví dụ ta có bảng dữ liệu số lượng và đơn giá bán như hình dưới. Yêu cầu tính tổng doanh thu của tất cả mặt hàng.


![](https://hddtvn.com/wp-content/uploads/2021/01/kknxxqJ.png)


Như bình thường, ta sẽ lấy cột B nhân với cột C để ra doanh thu như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/GSOBadZ.png)


Sau đó dùng hàm SUM để cộng tổng cột doanh thu.


![](https://hddtvn.com/wp-content/uploads/2021/01/9QyFoby.png)


Nhưng với hàm SUMPRODUCT, ta sẽ không cần làm nhiều bước như trên nữa. Chỉ bằng cách sử dụng hàm SUMPRODUCT ta có thể nhân doanh thu của tất cả mặt hàng sau đó cộng tộng ngay lập tức. Áp dụng cấu trúc hàm bên trên ta có công thức tính tổng doanh thu như sau:


**=SUMPRODUCT(B2:B7;C2:C7)**


![](https://hddtvn.com/wp-content/uploads/2021/01/bSEcjh1.png)


Ta có thể thấy, kết quả của hàm SUMPRODUCT giống như kết quả trước đó. Chỉ với một hàm duy nhất ta đã có thể tiết kiệm rất nhiều thời gian tính toán và giảm thiểu sai sót.


*Chúc các bạn thành công!*


moreHàm SUMPRODUCT là sự kết hợp của hàm SUM (tính tổng) và hàm PRODUCT (tính…

