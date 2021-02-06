Cách sử dụng linh hoạt hàm SUBTOTAL trong Excel
===============================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-linh-hoat-ham-subtotal-trong-excel/)Hàm SUBTOTAL là hàm rất linh hoạt trong Excel. Nó có thể tính toán hoặc thực hiện các phép tính như đếm số ô, tính trung bình, tìm giá trị lớn, nhỏ nhất… Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm SUBTOTAL hiệu quả nhất. 1. Cú pháp của hàm …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm SUBTOTAL là hàm rất linh hoạt trong Excel. Nó có thể tính toán hoặc thực hiện các phép tính như đếm số ô, tính trung bình, tìm giá trị lớn, nhỏ nhất… Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm SUBTOTAL hiệu quả nhất.**


### 1. Cú pháp của hàm SUBTOTAL


**SUBTOTAL(function\_num; ref1; [ref2];…)**


Trong đó:


**Function\_num:** Con số quy định hàm được sử dụng để tính toán, tương ứng với các hàm:




* Hàm AVERAGE (1 hoặc 101): Tính trung bình các con số

* Hàm COUNT (2 hoặc 102): Đếm số ô chứa giá trị số

* Hàm COUNTA (3 hay 103): Đếm số ô không trống

* Hàm MAX (4 hay 104): Tìm giá trị lớn nhất

* Hàm MIN (5 hay 105): Tìm giá trị nhỏ nhất

* Hàm PRODUCT (6 hay 106): Nhân các ô

* Hàm STDEV (7 hay 107): Tính độ lệch chuẩn mẫu dựa trên mẫu

* Hàm STDEVP (8 hay 108): Tính độ lệch chuẩn dựa trên toàn bộ số

* Hàm SUM (9 hay 109): Cộng các số

* Hàm VAR (10 hay 110): Ước tính độ dao động dựa trên mẫu

* Hàm VARP (11 hay 111): Ước tính độ dao động dựa trên toàn bộ số



**Ref1, Ref2, …:** 1 hoặc nhiều ô, hoặc dãy ô để tính tổng phụ. Cần phải có Ref 1, từ Ref 2 đến 254 là tuỳ chọn


***Chú ý:**Hàm sẽ bỏ qua không tính toán các hàng bị ẩn bởi lệnh Filter. Do đó nếu vùng dữ liệu đã sử dụng Filter thì khi sử dụng hàm Subtotal sẽ cho ra kết quả không đúng.


Bạn không cần phải nhớ hết các con số chức năng. Ngay khi bạn nhập hàm SUBTOTAL vào 1 ô hoặc trên thanh công thức, Excel sẽ đưa ra danh sách các con số để bạn chọn.


[![](https://hddtvn.com/wp-content/uploads/2021/01/OIzibDK.png)](https://hddtvn.com/wp-content/uploads/2021/01/OIzibDK.png)


### 2. Cách sử dụng hàm SUBTOTAL


#### Tính số tiền hàng trung bình


Chúng ta sử dụng hàm Average tương ứng với function\_num = 1


Công thức tại ô F12: **=SUBTOTAL(1; F2: F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/9mhjsBc.png)


#### Tính số lượng mặt hàng đã bán được


Chúng ta sử dụng hàm đếm Count tương ứng với Function\_num = 2.


Công thức tại ô F13: **=SUBTOTAL(2, F2: F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/z9sgiFj.png)


#### Tính số tiền hàng của mặt hàng bán được nhiều nhất


Chúng ta sử dụng hàm Max tương ứng với Function\_num = 4


Công thức tại ô F14: **=SUBTOTAL(4, F2: F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/8pAuUm6.png)


#### Tính số tiền hàng của mặt hàng bán được ít nhất


Chúng ta sử dụng hàm Min tương ứng với Function\_num = 5


Công thức tại ô F15: **=SUBTOTAL(5, F2: F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/TUPYVpp.png)


#### Tính tổng số tiền đã bán được:


Chúng ta sử dụng hàm tính tổng là hàm Sum tương ứng với Function = 9


Công thức tại ô F16: **=SUBTOTAL(9, F2: F11)**


![](https://hddtvn.com/wp-content/uploads/2021/01/C6oYknJ.png)


*Như vậy bài viết trên đã hướng dẫn các bạn sử dụng hàm SUBTOTAL trong Excel. Hàm SUBTOTAL có thể được sử dụng linh hoạt trong rất nhiều trường hợp. Muốn sử dụng tốt được hàm này thì chúng ta cần thực hành thật nhiều trong thực tiễn. Chúc các bạn thành công!*


moreHàm SUBTOTAL là hàm rất linh hoạt trong Excel. Nó có thể tính toán hoặc…

