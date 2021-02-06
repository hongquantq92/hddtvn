Cách sử dụng hàm DDB để tính khấu hao TSCĐ bằng phương pháp số dư giảm dần kép
==============================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-ddb-de-tinh-khau-hao-tscd-bang-phuong-phap-so-du-giam-dan-kep/)Trong bài viết sau đây, hddtvn sẽ hướng dẫn các bạn cách sử dụng hàm DDB để tính khấu hao TSCĐ bằng phương pháp số dư giảm dần kép. 1. Cấu trúc hàm DDB Cú pháp hàm: =DDB(cost, salvage, life, period, [factor]) Trong đó: Cost: đối số bắt buộc, là chi phí ban đầu của …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết sau đây, hddtvn sẽ hướng dẫn các bạn cách sử dụng hàm DDB để tính khấu hao TSCĐ bằng phương pháp số dư giảm dần kép.**


### 1. Cấu trúc hàm DDB


Cú pháp hàm: **=DDB(cost, salvage, life, period, [factor])**


Trong đó:




* **Cost**: đối số bắt buộc, là chi phí ban đầu của tài sản.

* **Salvage**: đối số bắt buộc, là giá trị sau khi khấu hao (đôi lúc được gọi là giá trị thu hồi của tài sản). Giá trị này có thể bằng 0.

* **Life**: đối số bắt buộc, là số kỳ khấu hao tài sản (đôi khi được gọi là tuổi thọ hữu ích của tài sản).

* **Period**: đối số bắt buộc, là số kỳ mà bạn muốn tính khấu hao. Kỳ khấu hao phải dùng cùng đơn vị với tuổi thọ.

* **Factor**: đối số tùy chọn, là tỷ lệ để giảm dần số dư. Nếu bỏ qua đối số **factor**, nó được giả định bằng 2 (phương pháp số dư giảm kép).



*Lưu ý:*




* Tất cả đối số đều phải là số dương

* Phương pháp số dư giảm dần kép tính khấu hao theo tỷ suất tăng dần. Khấu hao lớn nhất trong kỳ thứ nhất và giảm trong các kỳ kế tiếp. Hàm DDB dùng công thức sau đây để tính toán khấu hao trong một kỳ:  

*Min( (chi phí – tổng khấu hao từ những kỳ trước) * (hệ số/tuổi thọ); (chi phí – thu hồi – tổng khấu hao từ những kỳ trước) )*



### 2. Cách sử dụng hàm DDB


Ví dụ ta có một tài sản có chi phí ban đầu bỏ ra là 100.000.000. Giá trị thu hồi bằng 0 và có thời gian khấu hao là 5 năm.


![](https://hddtvn.com/wp-content/uploads/2021/01/GWpyvQJ.png)


Áp dụng cấu trúc hàm DDB như trên, ta có công thức tính khấu hao trong năm thứ 1 như sau:


**=DDB(B1;B2;B3;1)**


![](https://hddtvn.com/wp-content/uploads/2021/01/qPNZLrw.png)


Trong trường hợp sử dụng factor là 1,5. Ta có công thức tính khấu hao cho năm thứ 1 như sau:


**=DDB(B1;B2;B3;1;1,5)**


![](https://hddtvn.com/wp-content/uploads/2021/01/5Psn07V.png)


Để tính khấu hao cho tháng, ta nhân thời gian khấu hao (năm) với 12. Ví dụ để tính khấu hao trong tháng thứ 10, ta có công thức như sau:


**=DDB(B1;B2;B3*12;10)**


![](https://hddtvn.com/wp-content/uploads/2021/01/hlJW7pA.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DDB để tính khấu hao TSCĐ trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong bài viết sau đây, Ketoan.vn sẽ hướng dẫn các bạn cách sử dụng hàm…

