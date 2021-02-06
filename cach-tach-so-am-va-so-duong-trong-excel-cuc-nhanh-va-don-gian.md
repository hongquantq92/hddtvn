Cách tách số âm và số dương trong Excel cực nhanh và đơn giản
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-tach-so-am-va-so-duong-trong-excel-cuc-nhanh-va-don-gian/)Thay vì phải ngồi dò từng số thì với hàm IF bạn sẽ dễ dàng tìm nhanh được số âm và số dương. Nhờ đó giúp thao tác quản lý số liệu trở nên thuận tiện hơn. Bài viết sẽ hướng dẫn chi tiết cách sử dụng hàm IF để tách số âm và số …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Thay vì phải ngồi dò từng số thì với hàm IF bạn sẽ dễ dàng tìm nhanh được số âm và số dương. Nhờ đó giúp thao tác quản lý số liệu trở nên thuận tiện hơn. Bài viết sẽ hướng dẫn chi tiết cách sử dụng hàm IF để tách số âm và số dương trong Excel.**


![](https://hddtvn.com/wp-content/uploads/2021/01/tach-so-am-so-duong.png)


### 1. Cấu trúc hàm IF


Cú pháp hàm: **=IF (logical\_test, [value\_if\_true], [value\_if\_false])**


*Trong đó:*




* **Logical\_test**: Đối số bắt buộc, là một giá trị hay biểu thức logic có giá trị TRUE (đúng) hoặc FALSE (sai).

* **Value\_if\_true**: Đối số tùy chọn, là giá trị mà hàm sẽ trả về nếu biểu thức logic cho giá trị TRUE hay nói cách khác là điều kiện thỏa mãn.

* **Value\_if\_false**: Đối số tùy chọn, là giá trị mà hàm sẽ trả về nếu biểu thức logic cho giá trị FALSE hay nói cách khác là điều kiện không thỏa mãn.



### 2. Cách sử dụng hàm IF để tách số âm và số dương


Ví dụ ta có bảng dữ liệu như hình dưới. Yêu cầu cần tách riêng giá trị âm và giá trị dương. Để thực hiện, các bạn hãy làm theo các bước sau.


![](https://hddtvn.com/wp-content/uploads/2021/01/cGl6jTN.png)


**Bước 1:** Đầu tiên, các bạn áp dụng công thức hàm IF để nhập công thức tại ô đầu tiên của cột **Giá trị âm** như sau:


**=IF(D2<0;D2;””)**


Công thức trên có nghĩa nếu giá trị của ô D2 nhỏ hơn 0 thì kết quả sẽ trả về giá trị của ô D2, nếu không thì sẽ trả về ô trống.


![](https://hddtvn.com/wp-content/uploads/2021/01/f2g9mbj.png)


**Bước 2:** Tiếp theo, các bạn sao chép công thức cho tất cả ô còn lại. Ta sẽ thu được kết quả là các số âm đã được tách riêng tại cột **Giá trị âm**.


![](https://hddtvn.com/wp-content/uploads/2021/01/uQDU3Z8.png)


**Bước 3:** Tiếp theo, các bạn nhập công thức sau tại ô đầu tiên của cột Giá trị dương:


**=IF(D2>=0;D2;””)**


Công thức trên có nghĩa nếu giá trị của ô D2 lớn hơn hoặc bằng 0 thì kết quả trả về giá trị của ô D2, nếu không thì sẽ trả về ô trống.


![](https://hddtvn.com/wp-content/uploads/2021/01/I3UY7cV.png)


**Bước 4:** Các bạn sao chép công thức cho tất cả ô còn lại. Ta sẽ thu được kết quả là các số dương đã được tách riêng tại cột **Giá trị dương**. Chỉ cần vài bước đơn giản như vậy, ta đã sử dụng hàm IF để tách riêng số âm và số dương một cách dễ dàng.


![](https://hddtvn.com/wp-content/uploads/2021/01/fZQfV2O.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm IF để tách số âm và số dương trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreThay vì phải ngồi dò từng số thì với hàm IF bạn sẽ dễ dàng…

