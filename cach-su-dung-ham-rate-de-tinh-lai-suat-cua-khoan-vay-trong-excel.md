Cách sử dụng hàm RATE để tính lãi suất của khoản vay trong Excel
================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-rate-de-tinh-lai-suat-cua-khoan-vay-trong-excel/)Trong bài viết này, hddtvn sẽ cùng bạn đọc tìm hiểu về cách sử dụng hàm RATE để tính lãi suất của khoản vay trong Excel. 1. Cấu trúc hàm RATE Cú pháp hàm: =RATE(nper, pmt, pv, [fv], [type], [guess]) Trong đó: Nper: đối số bắt buộc, là kỳ hạn của khoản vay. Pmt: đối …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn sẽ cùng bạn đọc tìm hiểu về cách sử dụng hàm RATE để tính lãi suất của khoản vay trong Excel.**


### 1. Cấu trúc hàm RATE


Cú pháp hàm: **=RATE(nper, pmt, pv, [fv], [type], [guess])**


*Trong đó:*




* **Nper**: đối số bắt buộc, là kỳ hạn của khoản vay.

* **Pmt**: đối số bắt buộc, là số tiền thanh toán mỗi kỳ.

* **Pv**: đối số bắt buộc, là số tiền vay.

* **Fv**: đối số tùy chọn, là số dư tiền mặt bạn muốn thu được sau khi thực hiện khoản thanh toán cuối cùng.

* **Type**: đối số tùy chọn, là số 0 hoặc 1 chỉ rõ thời điểm thanh toán đến hạn.

* **Guess**: đối số tùy chọn, là ước đoán của bạn về lãi suất.



*Lưu ý:*




* Đối số **pmt** bao gồm cả tiền gốc và lãi thanh toán mỗi kỳ. Nếu **pmt** được bỏ qua, bạn phải đưa vào đối số **fv**.

* Nếu **fv** được bỏ qua, thì nó được giả định là 0.

* Nếu **type** là 0 hoặc bỏ qua, có nghĩa là thời điểm thanh toán vào cuối mỗi kỳ.

* Nếu **type** là 1, có nghĩa là thời điểm thanh toán vào đầu mỗi kỳ.

* Nếu **guess** bị bỏ qua số đoán, nó được giả định là 10 phần trăm.



### 2. Cách sử dụng hàm RATE


Ví dụ bạn vay ngân hàng một khoản 100tr. Hàng tháng bạn phải trả gốc và lãi 2tr trong vòng 6 năm. Yêu cầu tính toán lãi suất của khoản vay.


![](https://hddtvn.com/wp-content/uploads/2021/01/danQ6BV.png)


Áp dụng cấu trúc hàm như trên ta có công thức tính lãi suất hàng tháng như sau:


**=RATE(B5*12;-B4;B3)**


Ta cần lấy B5 nhân với 12 do thời hạn vay ở đây là năm còn số tiền thanh toán là hàng tháng.


![](https://hddtvn.com/wp-content/uploads/2021/01/OoLu1dc.png)


Để tính lãi suất theo năm thì ta chỉ cần lãi suất theo tháng nhân với 12. Hoặc sử dụng hàm như sau:


**=RATE(B5;-B4*12;B3)**


![](https://hddtvn.com/wp-content/uploads/2021/01/crV66MS.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm RATE để tính lãi suất của khoản vay trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong bài viết này, Ketoan.vn sẽ cùng bạn đọc tìm hiểu về cách sử dụng…

