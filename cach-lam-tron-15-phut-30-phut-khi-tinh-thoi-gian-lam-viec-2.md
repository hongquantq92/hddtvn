Cách làm tròn 15 phút/30 phút khi tính thời gian làm việc
=========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-lam-tron-15-phut-30-phut-khi-tinh-thoi-gian-lam-viec-2/)Với những bảng thống kê thời gian (như thời gian làm việc, thời gian tăng ca) bạn sẽ cần phải làm tròn giá trị thời gian về 15 phút hoặc 30 phút. Việc làm tròn như vậy sẽ có ích khi phải tổng kết, tính toán thời gian trong bảng lương nhân viên. Bài viết …
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Với những bảng thống kê thời gian (như thời gian làm việc, thời gian tăng ca) bạn sẽ cần phải làm tròn giá trị thời gian về 15 phút hoặc 30 phút. Việc làm tròn như vậy sẽ có ích khi phải tổng kết, tính toán thời gian trong bảng lương nhân viên. Bài viết sẽ chia sẻ với các bạn cách làm tròn thời gian 15 phút hoặc 30 phút trong Excel nhé.**


![Cách làm tròn thời gian 15 phút hoặc 30 phút trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/lam-tron-thoi-gian.png)


### 1. Cách làm tròn thời gian đến 15 phút


Ví dụ ta có bảng thời gian như hình dưới. Yêu cầu làm tròn thời gian đến 15 phút.


![](https://hddtvn.com/wp-content/uploads/2021/01/18-5.png)


Để làm tròn 15 phút, ta sẽ kết hợp sử dụng hàm FLOOR và hàm TIME. Cấu trúc của 2 hàm này như sau:


#### Cấu trúc hàm FLOOR


Cú pháp hàm: **=FLOOR(number; significance)**


*Trong đó:*




* **Number**: đối số bắt buộc, là số mà bạn muốn làm tròn.

* **Significance**: đối số bắt buộc, là bội số mà bạn muốn làm tròn đến.



#### Cấu trúc hàm TIME


Cú pháp hàm: **=TIME(hour; minute; second)**


*Trong đó:*




* **Hour**: đối số bắt buộc, là một số từ 0 đến 32767 thể hiện giờ.

* **Minute**: đối số bắt buộc, là một số từ 0 đến 32767 thể hiện phút.

* **Second**: đối số bắt buộc, là một số từ 0 đến 32767 thể hiện giây.



Áp dụng cấu trúc của 2 hàm bên trên ta có công thức làm tròn tới 15 phút tại ô B2 như sau:


**=FLOOR(A2;TIME(0;15;0))**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/19-5.png)


### 2. Làm tròn thời gian đến 30 phút


Cũng với ví dụ trên, nhưng nếu cần làm tròn thời gian 30 phút, ta sẽ kết hợp sử dụng hàm MOD và hàm INT. Cấu trúc của 2 hàm này như sau:


#### Cấu trúc hàm MOD


Cú pháp hàm: **=MOD (number; divisor)**


*Trong đó:*




* Number: đối số bắt buộc, là số bị chia.

* Divisor: đối số bắt buộc, là số chia.



#### Cấu trúc hàm INT


Cú pháp hàm: **=Int(Number)**


*Trong đó:*




* **Number**: đối số bắt buộc, là số thực hoặc là phép chia có số dư.



Áp dụng cấu trúc của 2 hàm bên trên ta có công thức làm tròn thời gian 30 phút tại ô C2 như sau:


**=MOD(INT(A2/(1/48))*(1/48);1)**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/20-3.png)


*Như vậy, bài viết trên đã giới thiệu đến các bạn cách làm tròn thời gian đến 15 phút hoặc 30 phút trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


#### 


moreVới những bảng thống kê thời gian (như thời gian làm việc, thời gian tăng…

