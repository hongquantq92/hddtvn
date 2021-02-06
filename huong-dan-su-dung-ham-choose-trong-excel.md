Hướng dẫn sử dụng hàm CHOOSE trong Excel
========================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-su-dung-ham-choose-trong-excel/)Hàm Choose trong Excel là hàm tìm kiếm 1 giá trị trong 1 chuỗi giá trị. Hàm thường được dùng để tính toán và tìm kiếm trong bảng dữ liệu Excel, khi có thể kết hợp với nhiều hàm Excel khác như hàm SUM trong Excel chẳng hạn. Khi kết hợp hàm Choose thì việc …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm Choose trong Excel là hàm tìm kiếm 1 giá trị trong 1 chuỗi giá trị. Hàm thường được dùng để tính toán và tìm kiếm trong bảng dữ liệu Excel, khi có thể kết hợp với nhiều hàm Excel khác như hàm SUM trong Excel chẳng hạn. Khi kết hợp hàm Choose thì việc tính toán sẽ đơn giản hơn. Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm CHOOSE trong Excel.**


![](https://hddtvn.com/wp-content/uploads/2021/01/choose.png)


### 1. Cấu trúc hàm CHOOSE


Cú pháp hàm: **=CHOOSE(index\_num; value1; [value2]; …)**


*Trong đó:*




* **Index\_num**: Đối số bắt buộc, là vị trí của dữ liệu trả về. **Index\_num** phải là số từ 1 đến 254 hoặc công thức hay tham chiếu đến ô chứa số từ 1 đến 254.





* **Value1**, **value2**, **…** : **Value 1** là bắt buộc, các giá trị tiếp theo là tùy chọn. Các đối số giá trị từ 1 đến 254 mà từ đó CHOOSE chọn giá trị hay hành động để thực hiện dựa trên **index\_num**. Các đối số có thể là số, tham chiếu ô, tên xác định, công thức, hàm hay văn bản.



*Chú ý:*




* Nếu **index\_num** là 1 thì hàm trả về value1; nếu nó là 2 thì hàm trả về value2; v.v..

* Nếu **index\_num** nhỏ hơn 1 hoặc lớn hơn số của giá trị cuối cùng trong danh sách, hàm sẽ trả về giá trị lỗi #VALUE!.

* Nếu **index\_num** không phải là số nguyên hàm trả về giá trị lỗi #Value!

* Nếu **index\_num** là phân số, nó bị cắt cụt đến số nguyên thấp nhất trước khi được dùng.

* Nếu **index\_num** là mảng, mọi giá trị được đánh giá khi đánh giá hàm CHOOSE.

* Các đối số giá trị cho hàm CHOOSE có thể là tham chiếu phạm vi cũng như giá trị đơn lẻ.



### 2. Cách sử dụng hàm CHOOSE


Ví dụ ta có bảng dữ liệu sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/fq23x8j.png)


Nếu cần lấy mã hàng thứ 3, ta có công thức: **=CHOOSE(3;B2;B3;B4;B5;B6;B7;B8;B9;B10)**


Hoặc **=CHOOSE(A4;B2;B3;B4;B5;B6;B7;B8;B9;B10)**


Hoặc ta cũng có thể viết thẳng mã hàng vào hàm như sau:


**=CHOOSE(3;0002;0005;0011;0013;0014;1001;1002;9003;9004)**


![](https://hddtvn.com/wp-content/uploads/2021/01/NtXTcbt.png)


Ta cũng có thể kết hợp hàm CHOOSE cùng với hàm SUM để tính tổng doanh thu như sau:


**=SUM(D2:CHOOSE(8;D3;D4;D5;D6;D7;D8;D9;D10))**


![](https://hddtvn.com/wp-content/uploads/2021/01/wDC4WbW.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm CHOOSE và cách kết hợp với hàm SUM trong Excel. Chúc các bạn thành công!*


moreHàm Choose trong Excel là hàm tìm kiếm 1 giá trị trong 1 chuỗi giá…

