Cách sử dụng hàm Average để tính trung bình cộng trong Excel
============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-average-de-tinh-trung-binh-cong-trong-excel/)Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Average để tính giá trị trung bình cộng cho một dãy số, một mảng hoặc tham chiếu chứa số một cách nhanh chóng và chính xác. 1. Công thức hàm AVERAGE Cú pháp hàm: =AVERAGE(number1;[number2];…) Trong đó Number1  Đối số bắt buộc. …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Average để tính giá trị trung bình cộng cho một dãy số, một mảng hoặc tham chiếu chứa số một cách nhanh chóng và chính xác.**


### 1. Công thức hàm AVERAGE


Cú pháp hàm: **=AVERAGE(number1;[number2];…)**


*Trong đó*




* **Number1**  Đối số bắt buộc. Số thứ nhất, ô tham chiếu, hoặc phạm vi muốn tính trung bình.

* **Number2,…** Đối số tùy chọn. Các số, ô tham chiếu hoặc phạm vi tiếp theo muốn tính trung bình, tối đa 255 đối số.



#### *Lưu ý*




* Đối số có thể là số, phạm vi hoặc ô tham chiếu có chứa số.

* Các giá trị logic và văn bản của các số nhập trực tiếp vào hàm sẽ không được tính.

* Nếu một ô tham chiếu hoặc phạm vi có chứa giá trị logic, văn bản hay ô trống thì những giá trị này sẽ bị bỏ qua. Tuy nhiên những ô có giá trị 0 sẽ được tính.

* Các đối số là văn bản hay giá trị lỗi không thể chuyển đổi thành số sẽ khiến xảy ra lỗi.

* Lưu ý tới sự khác biệt giữa ô trống và ô có chứa giá trị bằng không. Các ô trống không được tính, nhưng giá trị bằng không vẫn được tính.



Hàm AVERAGE tính ra số bình quân. Số bình quân thường được dùng với vai trò ước lượng xu hướng trung tâm. Xu hướng trung tâm là vị trí trung tâm của một nhóm số trong một phân bố thống kê. Ba cách đo lường thông dụng nhất về xu hướng trung tâm là:




* Trung bình, là trung bình số học, được tính bằng cách cộng một nhóm các số rồi chia cho số lượng các số. Ví dụ, trung bình của 2, 3, 3, 5, 7 và 10 là 30 chia cho 6, ra kết quả là 5.

* Trung vị, là số nằm ở giữa một nhóm các số; có nghĩa là, phân nửa các số có giá trị lớn hơn số trung vị, còn phân nửa các số có giá trị bé hơn số trung vị. Ví dụ, số trung vị của 2, 3, 3, 5, 7 và 10 là 4.

* Mode, là số xuất hiện nhiều nhất trong một nhóm các số. Ví dụ, mode của 2, 3, 3, 5, 7 và 10 là 3.



Với một phân phối đối xứng của một nhóm các số, ba cách đo lường về xu hướng trung tâm này đều là như nhau. Với một phân phối lệch của một nhóm các số, chúng có thể khác nhau.’


### 2. Cách sử dụng hàm AVERAGE để tính trung bình cộng


Ví dụ ta có bảng dữ liệu sau


![](https://hddtvn.com/wp-content/uploads/2021/01/X9qhAPU.png)


#### a. Tính doanh thu trung bình của các mặt hàng trong năm 2018


Công thức: **=AVERAGE(D2:D10)**


Ta thu được kết quả doanh thu trung bình năm 2018 của các mặt hàng


![](https://hddtvn.com/wp-content/uploads/2021/01/G39lDbg.png)


#### b. Tính doanh thu trung bình của tất cả các mặt hàng trong cả 3 năm


Công thức: **=AVERAGE(C2:E10)**


Ta được kết quả doanh thu trung bình tất cả mặt hàng trong 3 năm như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/yiN2WTv.png)


c. Tính doanh thu trung bình của Bia tươi ĐV đen trong 3 năm và doanh thu của Bia lon ĐV vàng năm 2017


Công thức: **=AVERAGE(C4:E4;C6)**


Ta thu được kết quả doanh thu trung bình của Bia tươi ĐV đen trong 3 năm và doanh thu của Bia lon ĐV vàng năm 2017:


![](https://hddtvn.com/wp-content/uploads/2021/01/fo3b7jI.png)


Hoặc ta cũng có thể thay ô tham chiếu doanh thu Bia lon ĐV vàng bằng cách viết thẳng số vào hàm như sau:


**=AVERAGE(C4:E4;982360300)**


Ta vẫn thu được kết quả tương tự


![](https://hddtvn.com/wp-content/uploads/2021/01/qpLcNtl.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm AVERAGE để tính trung bình cộng trong Excel. Chúc các bạn thành công!*


moreBài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm Average để…

