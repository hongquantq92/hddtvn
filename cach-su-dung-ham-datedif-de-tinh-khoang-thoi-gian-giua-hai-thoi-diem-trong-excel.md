Cách sử dụng hàm DATEDIF để tính khoảng thời gian giữa hai thời điểm trong Excel
================================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-datedif-de-tinh-khoang-thoi-gian-giua-hai-thoi-diem-trong-excel/)Excel hỗ trợ cho người dùng rất nhiều hàm để tính toán ngày tháng. Trong bài viết này, hddtvn sẽ giới thiệu với các bạn cách sử dụng hàm DATEDIF để tính khoảng thời gian giữa hai thời điểm trong Excel. 1. Cấu trúc hàm DATEDIF Cú pháp hàm: =DATEDIF(start\_date; end\_date; unit) Trong đó: Start\_date: đối …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel hỗ trợ cho người dùng rất nhiều hàm để tính toán ngày tháng. Trong bài viết này, hddtvn sẽ giới thiệu với các bạn cách sử dụng hàm DATEDIF để tính khoảng thời gian giữa hai thời điểm trong Excel.**


### 1. Cấu trúc hàm DATEDIF


Cú pháp hàm: **=DATEDIF(start\_date; end\_date; unit)**


*Trong đó:*




* **Start\_date**: đối số bắt buộc, là ngày tháng đầu tiên hoặc ngày bắt đầu của một khoảng thời gian.

* **End\_date**: đối số bắt buộc, là ngày tháng đầu tiên hoặc ngày bắt đầu của một khoảng thời gian.

* **Unit**: đối số bắt buộc, là loại thông tin mà bạn muốn trả về.



*Các loại của đối số Unit:*




* **Y**: Số năm hoàn tất trong khoảng thời gian.

* **M**: Số tháng hoàn tất trong khoảng thời gian.

* **D**: Số ngày trong khoảng thời gian.

* **MD**: Khoảng thời gian giữa các ngày trong **start\_date** và **end\_date**. Bỏ qua tháng và năm.

* **YM**: Khoảng thời gian giữa các tháng trong **start\_date** và **end\_date**. Bỏ qua ngày và năm.

* **Yd**: Khoảng thời gian giữa các ngày trong **start\_date** và **end\_date**. Bỏ qua năm.



*Lưu ý:* Nếu **start\_date** lớn hơn **end\_date** thì hàm sẽ trả về lỗi #NUM!.


### 2. Cách sử dụng hàm DATEDIF


Ví dụ ta cần tính khoảng thời gian giữa 2 thời điểm 4/5/2015 và 11/12/2019 theo các cách khác nhau bởi đối số Unit như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/7BUKNzE.png)


Như ta có thể thấy, hàm sẽ tính theo các cách khác nhau tùy vào đối số Unit:




* **Y**: Số năm giữa hai thời điểm 4/5/2015 và 11/12/2019.

* **M**: Số tháng giữa hai thời điểm 4/5/2015 và 11/12/2019.

* **D**: Số ngày giữa hai thời điểm 4/5/2015 và 11/12/2019.

* **MD**: Số ngày giữa ngày 4 và 11.

* **YM**: Số tháng giữa tháng 5 và 12.

* **Yd**: Số ngày giữa hai thời điểm 4/5 và 11/12.



*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DATEDIF trong Excel. Chúc các bạn thành công!*


moreExcel hỗ trợ cho người dùng rất nhiều hàm để tính toán ngày tháng. Trong…

