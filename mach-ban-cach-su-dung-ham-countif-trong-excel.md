Mách bạn cách sử dụng hàm COUNTIF trong Excel
=============================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/mach-ban-cach-su-dung-ham-countif-trong-excel/)Hàm COUNTIF sử dụng để đếm các ô chứa giá trị thỏa mãn một điều kiện cho trước. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm COUNTIF trong Excel. 1. Cấu trúc của hàm COUNTIF Cú pháp hàm: =COUNTIF(Range;Criteria) Trong đó: Range: là dãy dữ liệu chưa các ô mà …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm COUNTIF sử dụng để đếm các ô chứa giá trị thỏa mãn một điều kiện cho trước. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm COUNTIF trong Excel.**


### 1. Cấu trúc của hàm COUNTIF


Cú pháp hàm: **=COUNTIF(Range;Criteria)**


Trong đó:




* **Range**: là dãy dữ liệu chưa các ô mà bạn muốn đếm

* **Criteria**: là điều kiện để một ô được đếm



*Lưu ý*: Hàm COUNTIF chỉ áp dụng với 1 điều kiện.


### 2. Cách sử dụng hàm COUNTIF


Giả sử ta có danh sách học sinh như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/7JbeKUr.png)


#### a. Yêu cầu đếm số học sinh Nam.


Ta có công thức: **=COUNTIF(C2:C11;C2)**


Trong đó:




* **C2:C11** là mảng chứa các ô giới tính của học sinh

* **C2** là ô chứa từ Nam mà ta để làm điều kiện đếm



![](https://hddtvn.com/wp-content/uploads/2021/01/r7gk3zN.png)


Ta cũng có thể thay điều kiện **C2** thành **“Nam”**. Lúc này ta có công thức đếm: **=COUNTIF(C2:C11;”Nam”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/9uxzZeq.png)


#### b. Yêu cầu đếm số học sinh có điểm thi toán nhỏ hơn 5


Ta có công thức đếm: **=COUNTIF(D2:D11;”<5″)**


Trong đó:




* **D2:D11** là mảng chứa các ô điểm thi toán của học sinh

* **<5** là điều kiện đếm



![](https://hddtvn.com/wp-content/uploads/2021/01/9mCUBv9.png)


Hoặc ta có thể sử dụng ký tự **&** để so sánh với 1 ô chứa số 5: **=COUNTIF(D2:D11;”<“&D8)**


![](https://hddtvn.com/wp-content/uploads/2021/01/vt6fzJj.png)


Công thức cũng tương tự với các điều kiện đếm số lớn hơn, lớn hơn hoặc bằng, nhỏ hơn hoặc bằng.


#### c. Yêu cầu đếm số học sinh có tên bắt đầu bằng chữ H


Ta có thể dùng dấu * để thay các ký tự khác trong ô điều kiện. Lúc này ta có công thức đếm số học sinh có tên bắt đầu bằng chữ H:


**=COUNTIF(B2:B11;”H*”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/Pdl9V9E.png)


*Như vậy, bài viết trên đã hướng dẫn và cho ví dụ về các cách sử dụng hàm COUNTIF trong Excel. Chúc các bạn thành công!*


moreHàm COUNTIF sử dụng để đếm các ô chứa giá trị thỏa mãn một điều kiện cho trước. Bài viết sau đây sẽ hướng…

