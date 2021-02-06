Cách sử dụng hàm AMORLINIC để tính khấu hao kỳ hạn trong Excel
==============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-amorlinic-de-tinh-khau-hao-ky-han-trong-excel/)Excel cung cấp cho người dùng rất nhiều hàm để tính khấu hao tài sản theo nhiều phương pháp khác nhau. Trong bài viết này, hddtvn sẽ giới thiệu cách sử dụng hàm AMORLINIC trả về khấu hao cho mỗi kỳ kế toán bằng cách dùng hệ số khấu hao. Ví dụ cụ thể trong …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Excel cung cấp cho người dùng rất nhiều hàm để tính khấu hao tài sản theo nhiều phương pháp khác nhau. Trong bài viết này, hddtvn sẽ giới thiệu cách sử dụng hàm AMORLINIC trả về khấu hao cho mỗi kỳ kế toán bằng cách dùng hệ số khấu hao. Ví dụ cụ thể trong bài viết sẽ giúp bạn dễ dàng hình dung và nhanh biết cách sử dụng hàm này.**


### 1. Cấu trúc hàm AMORLINIC


Cú pháp hàm: **=AMORLINC(cost; date\_purchased; first\_period; salvage; period; rate; [basis])**


*Trong đó:*




* **Cost**: đối số bắt buộc, là chi phí của tài sản.

* **Date\_purchased**: đối số bắt buộc, là ngày mua tài sản.

* **First\_period**: đối số bắt buộc, là ngày kết thúc của kỳ thứ nhất.

* **Salvage**: đối số bắt buộc, là giá trị thu hồi của tài sản.

* **Period:** đối số bắt buộc, là số kỳ tính khấu hao.

* **Rate:** đối số bắt buộc, là tỷ lệ khấu hao.

* **Basis**: đối số tùy chọn, là cơ sở năm được dùng.  

– **Basis** = 0 hoặc bỏ qua: Một tháng có 30 ngày/Một năm có 360 ngày (theo tiêu chuẩn Bắc Mỹ).  

– **Basis** = 1: Số ngày thực tế của mỗi tháng / Số ngày thực tế của mỗi năm.  

– **Basis** = 2: Số ngày thực tế của mỗi tháng / Một năm có 360 ngày.  

– **Basis** = 3: Số ngày thực tế của mỗi tháng / Một năm có 365 ngày.  

– **Basis** = 4: Một tháng có 30 ngày / Một năm có 360 ngày (theo tiêu chuẩn Châu Âu).



### 2. Cách sử dụng hàm AMORLINIC


Ví dụ ta có bảng tính khấu hao tài sản cố định như hình dưới. Yêu cầu tính khấu hao kỳ hạn của từng tài sản bằng cách sử dụng hàm AMORLINIC.


![](https://hddtvn.com/wp-content/uploads/2021/01/8kXS46y.png)


Áp dụng công thức hàm bên trên, ta có công thức tính khấu hao kỳ hạn của Cổng chào như sau:


**=AMORLINC(C2;D2;E2;F2;G2;H2;3)**


Sao chép công thức cho các tài sản còn lại ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/vGiWE8r.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm AMORLINIC để tính khấu hao kỳ hạn của tài sản trong Excel. Chúc các bạn thành công!*


moreExcel cung cấp cho người dùng rất nhiều hàm để tính khấu hao tài sản…

