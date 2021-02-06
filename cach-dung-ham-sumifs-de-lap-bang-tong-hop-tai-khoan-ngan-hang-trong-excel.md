Cách dùng hàm SUMIFS để lập bảng tổng hợp tài khoản ngân hàng trong Excel
=========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-sumifs-de-lap-bang-tong-hop-tai-khoan-ngan-hang-trong-excel/)Trong bài viết này, hddtvn xin chia sẻ với các bạn cách sử dụng hàm SUMIFS để lập bảng tổng hợp tài khoản ngân hàng trong Excel. 1. Cấu trúc hàm SUMIFS Cú pháp hàm: =SUMIFS(sum\_range, criteria\_range1, criteria1, [criteria\_range2, criteria2], …) Trong đó: Sum\_range: đối số bắt buộc, là vùng cần tính tổng. Criteria\_range1: đối sốt …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn xin chia sẻ với các bạn cách sử dụng hàm SUMIFS để lập bảng tổng hợp tài khoản ngân hàng trong Excel.**


### 1. Cấu trúc hàm SUMIFS


Cú pháp hàm: **=SUMIFS(sum\_range, criteria\_range1, criteria1, [criteria\_range2, criteria2], …)**


*Trong đó:*




* **Sum\_range**: đối số bắt buộc, là vùng cần tính tổng.

* **Criteria\_range1**: đối sốt bắt buộc, là vùng chứa các ô điều kiện thứ 1.

* **Criteria**: đối số bắt buộc, là điều kiện thứ 1.

* **Criteria\_range2…**: đối số tùy chọn, là vùng chứa các ô điều kiện thứ 2 trở đi.

* **Criteria…**: đối số tùy chọn, là điều kiện thứ 2 trở đi.



*Lưu ý*: Nếu Criteria là điều kiện dạng văn bản hoặc có chứa các ký hiệu toán học thì đều phải được đặt trong dấu nháy kép (” “).


### 2. Cách sử dụng hàm SUMIFS để lập bảng tổng hợp tài khoản ngân hàng


Ví dụ ta có bảng tổng hợp tài khoản ngân hàng và sổ nhật ký chung như hình dưới. Yêu cầu cần tổng hợp phát sinh nợ và có của từng tải khoản để tính ra số dư cuối kỳ của từng tài khoản.


![](https://hddtvn.com/wp-content/uploads/2021/01/s2l7quB.png)


**Bước 1:** Vì nguồn dữ liệu là rất lớn nên việc đầu tiên ta cần làm đó là đặt tên cho mảng dữ liệu để tiện viết công thức. Các bạn có thể xem lại bài đặt tên cho vùng dữ liệu tại [đây](#)


Trong bài viết này mình đặt tên cho 4 vùng là Ngày, Nợ, Có, Tiền tương ứng với cột Ngày, Nợ, Có, Số tiền trong sheet NKC.


![](https://hddtvn.com/wp-content/uploads/2021/01/vcz7hKF.png)


**Bước 2:** Áp dụng cấu trúc hàm SUMIFS bên trên ta có công thức tính tổng phát sinh nợ của tài khoản 1121 từ ngày 01/01/2020 đến ngày 30/06/2020 như sau:


**=SUMIFS(Tiền;Nợ;A7;Ngày;”>=”&$G$3;Ngày;”<=”&$G$4)**


Sao chép công thức cho các tài khoản bên dưới ta sẽ thu được kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/WJbXt3A.png)


**Bước 3:** Ta có công thức tính tổng phát sinh có của tài khoản 1121 từ ngày 01/01/2020 đến ngày 30/06/2020 như sau:


**=SUMIFS(Tiền;Có;A7;Ngày;”>=”&$G$3;Ngày;”<=”&$G$4)**


Sao chép công thức cho các tài khoản bên dưới ta sẽ thu được kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/QbYN36D.png)


**Bước 4:** Sau khi tính được tổng phát sinh Nợ và tổng phát sinh Có, chúng ta có thể tính ra số dư cuối kỳ theo công thức:


**=D7+E7-F7**


Sao chép công thức cho các tài khoản bên dưới ta sẽ thu được bảng tổng hợp tài khoản ngân hàng hoàn chỉnh.


#### Các bạn có thể tải file tại đây


![](https://hddtvn.com/wp-content/uploads/2021/01/tlGPJPU.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm SUMIFS để lập bảng tổng hợp tài khoản ngân hàng. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong bài viết này, Ketoan.vn xin chia sẻ với các bạn cách sử dụng hàm…

