Cách sử dụng hàm CUMPRINC để tính số tiền gốc dồn tích trả cho khoản vay trong Excel
====================================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-cumprinc-de-tinh-so-tien-goc-don-tich-tra-cho-khoan-vay-trong-excel/)Hàm CUMPRINC là một trong những hàm tài chính thông dụng trong Excel. Hàm CUMPRINC thường được sử dụng để tính số tiền gốc dồn tích trả cho khoản vay từ kỳ đầu đến kỳ cuối. Trong bài viết này, hddtvn sẽ giới thiệu các bạn cách sử dụng hàm CUMPRINC trong Excel. 1. Cấu …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm CUMPRINC là một trong những hàm tài chính thông dụng trong Excel. Hàm CUMPRINC thường được sử dụng để tính số tiền gốc dồn tích trả cho khoản vay từ kỳ đầu đến kỳ cuối. Trong bài viết này, hddtvn sẽ giới thiệu các bạn cách sử dụng hàm CUMPRINC trong Excel.**


### 1. Cấu trúc hàm CUMPRINC


Cú pháp hàm: **=CUMPRINC(rate; nper; pv; start\_period; end\_period; type)**


*Trong đó:*




* **Rate**: đối số bắt buộc, là lãi suất của khoản vay.

* **Nper**: đối số bắt buộc, là tổng số kỳ thanh toán.

* **Pv**: đối số bắt buộc, là giá trị hiện tại của khoản vay.

* **Start\_period**: đối số bắt buộc, là kỳ đầu tiên. Các kỳ thanh toán được đánh số bắt đầu từ 1.

* **End\_period**: đối số bắt buộc, là kỳ cuối cùng.

* **Type**: đối số bắt buộc, là thời hạn thanh toán.



*Lưu ý:*




* Nếu **Type** là 0 có nghĩa là thanh toán vào cuối kỳ. Nếu là 1 có nghĩa là thanh toán vào đầu kỳ

* Nếu **type** là bất kỳ số nào khác 0 hoặc 1, hàm CUMPRINC trả về giá trị lỗi #NUM! .

* Nếu **rate** ≤ 0, **nper** ≤ 0 hoặc **pv** ≤ 0, hàm CUMPRINC trả về giá trị lỗi #NUM! .

* **Rate** và **nper** phải được xác định bằng đơn vị nhất quán.

* Nếu **start\_period** < 1, **end\_period** < 1 hoặc **start\_period** > **end\_period**, hàm CUMPRINC trả về giá trị lỗi #NUM!.



### 2. Cách sử dụng hàm CUMPRINC


Ví dụ bạn vay 100 triệu với lãi suất 11% trong vòng 5 năm. Yêu cầu tính số tiền phải trả trong tháng thứ nhất và số tiền phải trả trong năm thứ 3.


![](https://hddtvn.com/wp-content/uploads/2021/01/SiBJHin.png)


#### a. Tính số tiền phải trả trong tháng thứ nhất


Áp dụng cấu trúc hàm bên trên, ta có công thức tính số tiền phải trả trong tháng thứ nhất như sau:


**=CUMPRINC(B4/12;B5*12;B3;1;1;0)**


Trong hàm tính bên trên, do ta cần tính số tiền phải trả trong tháng nên đơn vị dùng cho Rate và Nper sẽ phải là tháng. Vì lãi suất là 11%/năm nên lãi suất tháng sẽ là 11%/12. Số kỳ thanh toán là 5 năm tương ứng với 60 tháng.


![](https://hddtvn.com/wp-content/uploads/2021/01/FVZPyQp.png)


#### b. Tính số tiền phải trả trong năm thứ 3


Áp dụng cấu trúc hàm bên trên, ta có công thức tính số tiền phải trả trong năm thứ 3 như sau:


**=CUMPRINC(B4;B5;B3;3;3;0)**


Do đơn vị tính lần này là năm nên các đối số Rate và Nper ta giữ nguyên.


![](https://hddtvn.com/wp-content/uploads/2021/01/jG2NZlu.png)


*Như vậy, bài viết trên đã giới thiệu đến các bạn cách sử dụng hàm CUMPRINC trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm CUMPRINC là một trong những hàm tài chính thông dụng trong Excel. Hàm CUMPRINC…

