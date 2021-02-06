Hướng dẫn cách tính thời gian làm thêm bằng Excel
=================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-tinh-thoi-gian-lam-them-bang-excel/)Tính thời gian làm thêm (thời gian OT) từ bảng chấm công bằng Excel là thao tác mà kế toán tiền lương phải làm hàng tháng. Tuy nhiên không phải ai cũng thành thạo việc kết hợp các hàm trong Excel để tính thời gian làm thêm. Mời bạn theo dõi bài viết sau để …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Tính thời gian làm thêm (thời gian OT) từ bảng chấm công bằng Excel là thao tác mà kế toán tiền lương phải làm hàng tháng. Tuy nhiên không phải ai cũng thành thạo việc kết hợp các hàm trong Excel để tính thời gian làm thêm. Mời bạn theo dõi bài viết sau để biết cách tính thời gian làm thêm bằng Excel nhé.**


![Hướng dẫn cách tính thời gian làm thêm bằng Excel](https://hddtvn.com/wp-content/uploads/2021/01/thoi-gian-lam-them.png)


Ví dụ ta có bảng chấm công như hình dưới. Theo quy định thì thời gian làm việc từ 17h trở đi sẽ là thời gian làm thêm. Yêu cầu tính thời gian làm thêm của các nhân viên trong hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/24.png)


### Cách 1: Lấy giờ ra trừ đi 17h


Cách đơn giản nhất để tính thời gian làm thêm đó chính là lấy giờ ra trừ đi 17h. Đầu tiên, các bạn cần nhập 17:00 vào một ô bất kỳ ngoài bảng chấm công. Sau đó các bạn nhập công thức tại ô E2: **=D2-$D$8**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả như hình dưới.


![](https://hddtvn.com/wp-content/uploads/2021/01/25-1.png)


Ngoài ra nếu các bạn không muốn nhập 17:00 ở một ô ngoài bảng thì ta có thể sử dụng hàm TIME thay thế như sau:


**=D2-TIME(17;0;0)**


Sao chép công thức cho tất cả các ô còn lại ta sẽ thu được kết quả như hình dưới.


t


### Cách 2: Kết hợp các hàm để tính thời gian làm thêm


Để sử dụng hàm Excel tính thời gian làm thêm, chúng ta cần kết hợp các hàm IF, HOUR, TIME. Các bạn nhập công thức sau đây vào ô F2:


**=IF(HOUR(D2)>=17;ROUND((D2-TIME(17;0;0))*24*60;0);0)**


Công thức trên có nghĩa là nếu giờ trong ô D2 lớn hơn 17 thì kết quả sẽ trả về số phút được làm tròn của thời gian trong ô D2 trừ đi 17h. Nếu giờ trong ô D2 nhỏ hơn 17 thì kết quả sẽ trả về là 0.


Sao chép công thức trên cho các ô còn lại thì ta sẽ thu được kết quả là tính được thời gian làm thêm của từng nhân viên theo số phút. Đối chiếu với thời gian làm thêm tính theo cách trên thì kết quả giống nhau chỉ khác định dạng giờ và phút.


![](https://hddtvn.com/wp-content/uploads/2021/01/YJDrigc.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tính thời gian làm thêm bằng Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTính thời gian làm thêm (thời gian OT) từ bảng chấm công bằng Excel là…

