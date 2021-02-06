Hướng dẫn dùng hàm SUMIFS để tính tổng theo nhiều điều kiện trong Excel
=======================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-dung-ham-sumifs-de-tinh-tong-theo-nhieu-dieu-kien-trong-excel/)Excel hỗ trợ cho người dùng rất nhiều hàm để tính tổng. Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng hàm SUMIFS để tính tổng theo nhiều điều kiện trong Excel. 1. Cấu trúc hàm SUMIFS Cú pháp hàm: =SUMIFS(sum\_range, criteria\_range1, criteria1, [criteria\_range2, criteria2], …) Trong đó: Sum\_range: đối …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel hỗ trợ cho người dùng rất nhiều hàm để tính tổng. Trong bài viết này, hddtvn sẽ giới thiệu đến các bạn cách sử dụng hàm SUMIFS để tính tổng theo nhiều điều kiện trong Excel.**


### 1. Cấu trúc hàm SUMIFS


Cú pháp hàm: **=SUMIFS(sum\_range, criteria\_range1, criteria1, [criteria\_range2, criteria2], …)**


*Trong đó:*




* **Sum\_range**: đối số bắt buộc, là vùng cần tính tổng.

* **Criteria\_range1**: đối sốt bắt buộc, là vùng chứa các ô điều kiện thứ 1.

* **Criteria**: đối số bắt buộc, là điều kiện thứ 1.

* **Criteria\_range2…**: đối số tùy chọn, là vùng chứa các ô điều kiện thứ 2 trở đi.

* **Criteria…**: đối số tùy chọn, là điều kiện thứ 2 trở đi.



*Lưu ý*: Nếu Criteria là điều kiện dạng văn bản hoặc có chứa các ký hiệu toán học thì đều phải được đặt trong dấu nháy kép (” “).


### 2. Cách sử dụng hàm SUMIFS


Giả sử ta có bảng doanh thu như hình dưới. Yêu cầu tính tổng doanh thu của các sản phẩm là Bột đá và có số lượng lớn hơn 30.000.


![](https://hddtvn.com/wp-content/uploads/2021/01/4Uno1Hv.png)


Áp dụng cấu trúc hàm bên trên ta có công thức tính tổng doanh thu các sản phẩm là bột đá có số lượng lớn hơn 30.000 như sau:


**=SUMIFS(F2:F11;D2:D11;D3;E2:E11;”>30000″)**


Ta có thể thấy 2 sản phẩm thỏa mãn điều kiện là Đá 0x4 và Đá mi bụi với doanh thu lần lượt là 9.154.583.427 và 5.036.227.924. Cộng doanh thu 2 sản phẩm này vào ta được kết quả của hàm là 14.190.811.351.


![](https://hddtvn.com/wp-content/uploads/2021/01/MgJoBpX.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm SUMIFS để tính tổng theo nhiều điều kiện trong Excel. Hy vọng bài viết trên sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreExcel hỗ trợ cho người dùng rất nhiều hàm để tính tổng. Trong bài viết…

