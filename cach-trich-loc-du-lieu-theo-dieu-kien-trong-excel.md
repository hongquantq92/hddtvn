Cách trích, lọc dữ liệu theo điều kiện trong Excel
==================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-trich-loc-du-lieu-theo-dieu-kien-trong-excel/)Trong quá trình làm việc trên Excel, nhiều lúc các bạn sẽ gặp các bảng tính chứa nhiều thông tin khác nhau. Lúc này việc trích, lọc dữ liệu sẽ giúp bạn dễ dàng tìm thấy dữ liệu thỏa mãn một hoặc một số điều kiện nào đó một cách dễ dàng, chính xác hơn. …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình làm việc trên Excel, nhiều lúc các bạn sẽ gặp các bảng tính chứa nhiều thông tin khác nhau. Lúc này việc trích, lọc dữ liệu sẽ giúp bạn dễ dàng tìm thấy dữ liệu thỏa mãn một hoặc một số điều kiện nào đó một cách dễ dàng, chính xác hơn. Bài viết sau đây sẽ hướng dẫn các bạn cách trích, lọc dữ liệu trong Excel.**


### 1. Sử dụng Filter để lọc dữ liệu theo điều kiện


Ví dụ: ta có bảng dữ liệu cần lọc những học sinh có cùng tổng điểm:


[![](https://hddtvn.com/wp-content/uploads/2021/01/JlQ0ntN.png)](https://hddtvn.com/wp-content/uploads/2021/01/JlQ0ntN.png)


Ta chỉ cần bôi đen vùng dữ liệu cần lọc => chọn thẻ **Data** => chọn biểu tượng **Fillter**


![](https://hddtvn.com/wp-content/uploads/2021/01/tmVWzck.png)


Lúc này bảng dữ liệu sẽ xuất hiện mũi tên trên hàng tiêu đề. Ví dụ muốn lọc những hoc sinh có điểm thi toán bằng 6.


Ta chọn mũi tên trong cột Điểm thi toán => tích bỏ chọn mục **Select All** => tích chọn giá trị 6 => **OK**


![](https://hddtvn.com/wp-content/uploads/2021/01/YGojF5Y.png)


Kết quả ta đã lọc những học sinh có cùng tổng điểm bằng 6


![](https://hddtvn.com/wp-content/uploads/2021/01/NwtW8DH.png)


Nếu bạn muốn lọc nhiều giá trị, ví dụ lọc học sinh có Điểm thi toán từ 6 đến 9. Ta chỉ cần tích chọn các giá từ 6 đến 9. Tuy nhiên cách này sẽ không khả thi với những trường hợp có quá nhiều giá trị.


![](https://hddtvn.com/wp-content/uploads/2021/01/3ktU4Hb.png)


Kết quả ta đã lọc được những học sinh có điểm thi toán từ 6 đến 9.


![](https://hddtvn.com/wp-content/uploads/2021/01/NwgIxLi.png)


### 2. Sử dụng Advanced Filter để trích lọc dữ liệu có điều kiện


#### a. Trích lọc với 1 điều kiện


Với lượng dữ liệu lớn bạn nên sử dụng cách này để trích lọc dữ liệu.


Ví dụ muốn trích lọc học sinh có Điểm thi toán lớn hơn 5. Danh sách trích lọc đặt sang ví khác để tiện cho quá trình in ấn.


Ta cần đặt điều kiện với tiêu đề cột và giá trị xác định điều kiện. Tên tiêu đề phải trùng với tiêu đề trong bảng cần trích lọc dữ liệu. Sau đó vào thẻ **Data** => chọn biểu tượng **Advanced**


![](https://hddtvn.com/wp-content/uploads/2021/01/XiqBqoR.png)


Hộp thoại **Advanced** xuất hiện, ta lựa chọn như sau:




* Tích chọn vào mục **Copy to another location**: để lựa chọn nội dung đã trích lọc sang vị trí mới.

* Mục **List range**: Kích chọn mũi tên để lựa chọn vùng dữ liệu chứa dữ liệu cần trích lọc.

* Mục **Cirteria range**: Lựa chọn điều kiện trích lọc dữ liệu

* Mục **Copy to**: Lựa chọn vị trí dán nội dung dữ liệu sau khi đã trích lọc được.



Cuối cùng kích chọn **OK**


![](https://hddtvn.com/wp-content/uploads/2021/01/77W1LcQ.png)


Dữ liệu đã được trích lọc sang vị trí mới mà bạn đã chọn


![](https://hddtvn.com/wp-content/uploads/2021/01/P40ssCc.png)


#### b. Trích lọc dữ liệu với nhiều điều kiện


Với việc trích lọc dữ liệu có chứa nhiều điều kiện bạn cần chú ý,các điều kiện cùng nằm trên 1 hàng và tên tiêu đề cột chứa điều kiện trích lọc phải trùng với tên tiêu đề cột trong bảng dữ liệu nguồn.


Ví dụ cần trích lọc học sinh có Điểm thi toán và điểm thi văn lớn hơn 5. Ta thực hiện tạo cột điều kiện như hình dưới sau đó vào thẻ **Data** => chọn biểu tượng **Advanced**


![](https://hddtvn.com/wp-content/uploads/2021/01/2aBvbN6.png)


Tương tự việc lọc dữ liệu với một điều kiện, với nhiều điều kiện trong mục **Cirteria range** bạn lựa chọn toàn bộ điều kiện đã tạo:


![](https://hddtvn.com/wp-content/uploads/2021/01/Rtc1nfb.png)


Kết quả bạn đã trích lọc dữ liệu với 2 điều kiện là điểm thi toán và điểm thi văn lớn hơn 5


![](https://hddtvn.com/wp-content/uploads/2021/01/LFV7tQB.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách để trích lọc dữ liệu theo điều kiện trong Excel. Chúc các bạn thành công!*


moreTrong quá trình làm việc trên Excel, nhiều lúc các bạn sẽ gặp các bảng…

