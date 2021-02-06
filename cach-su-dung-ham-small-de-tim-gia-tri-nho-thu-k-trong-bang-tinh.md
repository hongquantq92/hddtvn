Cách sử dụng hàm SMALL để tìm giá trị nhỏ thứ K trong bảng tính
===============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-small-de-tim-gia-tri-nho-thu-k-trong-bang-tinh/)Bạn chắc đã làm quen hàm Min và MinA để tìm giá trị nhỏ nhất của dãy số, của vùng dữ liệu? Vậy nếu muốn tìm ra những con số nhỏ thứ 2, thứ 3 hay thứ K trong dãy số thì làm thế nào? Bài viết dưới đây sẽ giúp bạn làm thống kê …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn chắc đã làm quen hàm Min và MinA để tìm giá trị nhỏ nhất của dãy số, của vùng dữ liệu? Vậy nếu muốn tìm ra những con số nhỏ thứ 2, thứ 3 hay thứ K trong dãy số thì làm thế nào? Bài viết dưới đây sẽ giúp bạn làm thống kê báo cáo nhanh hơn, hiệu quả hơn đó là sử dụng hàm SMALL.**


### 1. Cấu trúc hàm SMALL


Cú pháp hàm: **=SMALL(array; k)**


*Trong đó:*




* **Array**: đối số bắt buộc, là mảng hoặc phạm vi dữ liệu dạng số mà bạn muốn xác định giá trị nhỏ thứ k của nó.

* **K**: đối số bắt buộc, là vị trí (từ giá trị nhỏ nhất) trong mảng hoặc phạm vi dữ liệu cần trả về.



Lưu ý:




* Nếu **k** ≤ 0 hoặc nếu **k** vượt quá số lượng giá trị của **array**, hàm sẽ trả về giá trị lỗi #NUM! .

* Nếu **Array** trống, hàm sẽ trả về giá trị lỗi #NUM! .

* Nếu n là số lượng giá trị của array thì **=SMALL(array;1)** sẽ trả về giá trị nhỏ nhất và **=SMALL(array;n)** sẽ ́trả ́vệ̀ gíá tŕị lớn nhất.



### 2. Cách sử dụng hàm SMALL


Ví dụ ta có bảng điểm thi toán và văn của các học sinh như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/WyPeFR1.png)


Để tìm điểm thi văn nhỏ thứ 7 của học sinh, ta có công thức như sau:


**=SMALL(E2:E11;7)**


![](https://hddtvn.com/wp-content/uploads/2021/01/n0LuI8O.png)


Để tìm điểm thi toán nhỏ thứ 3 của học sinh, ta có công thức như sau:


**=SMALL(D2:D11;3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/hF4GJXo.png)


Ta có thể thấy có 3 học sinh đều có điểm thi toán là 4. Lúc này hàm SMALL sẽ không tính cả 3 học sinh đều hạng 3 mà sẽ là hạng 3,4,5. Như vậy nếu ta tìm học sinh có điểm thi toán nhỏ thứ 3, 4, 5 thì kết quả thu được sẽ đều là 4.


![](https://hddtvn.com/wp-content/uploads/2021/01/yi8y4CT.png)


Vì bảng trên chỉ có 10 học sinh nên nếu ta tìm điểm thi toán nhỏ thứ 11 trở lên thì hàm sẽ trả về giá trị lỗi #NUM!.


![](https://hddtvn.com/wp-content/uploads/2021/01/WdEmTXj.png)


*Hy vọng qua bài viết này, các bạn đã thành thạo được cách sử dụng hàm SMALL trong Excel. Chúc các bạn thành công!*


moreBạn chắc đã làm quen hàm Min và MinA để tìm giá trị nhỏ nhất…

