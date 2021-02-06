Hướng dẫn sử dụng các hàm COUNT, COUNTA, COUNTBLACK để đếm số ô trong Excel
===========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-su-dung-cac-ham-count-counta-countblack-de-dem-so-o-trong-excel/)Khi làm việc trên Excel, nhiều lúc bạn sẽ cần đếm các ô có giá trị, hoặc các ô trống. Trong Excel có rất nhiều hàm để thực hiện việc đếm ô tùy theo nhu cầu sử dụng. Bài viết dưới đây sẽ hướng dẫn cách bạn cách sử dụng hàm COUNT, COUNTA, COUNTBLANK để …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Khi làm việc trên Excel, nhiều lúc bạn sẽ cần đếm các ô có giá trị, hoặc các ô trống. Trong Excel có rất nhiều hàm để thực hiện việc đếm ô tùy theo nhu cầu sử dụng. Bài viết dưới đây sẽ hướng dẫn cách bạn cách sử dụng hàm COUNT, COUNTA, COUNTBLANK để đếm số ô trong Excel**


### 1. Sử dụng hàm COUNT để đếm các ô có số


Hàm COUNT chỉ đếm số lượng ô có chứa số.


Cú pháp hàm: **=COUNT(value1;[value2];…)**


*Trong đó:*




* **value1**: Bắt buộc, là các ô tham chiếu hoặc các mảng trong đó bạn muốn đếm ô.

* **value2,…**:Tùy chọn, tối đa 255 mục, là các ô tham chiếu hoặc các mảng trong đó bạn muốn đếm ô.



*Chú ý:*




* Các đối số có thể chứa hoặc tham chiếu đến các kiểu dữ liệu khác nhau nhưng chỉ các số mới được đếm.

* Các đối số là số, ngày hay trình bày dạng văn bản của số (ví dụ, số nằm trong dấu trích dẫn, chẳng hạn như “1”) sẽ được đếm.

* Các giá trị logic và trình bày số dạng văn bản mà bạn gõ trực tiếp vào danh sách các đối số sẽ được đếm.

* Các đối số là văn bản hay giá trị lỗi không thể chuyển đổi thành số sẽ không được đếm.

* Nếu đối số là mảng hay tham chiếu, chỉ các các số trong mảng hay tham chiếu đó mới được đếm. Các ô trống, giá trị logic, văn bản hoặc giá trị lỗi trong mảng hoặc tham chiếu sẽ không được đếm.



Ví dụ: Ta có danh sách điểm thi toán theo bảng dưới. Cần đếm số bạn đã có điểm thi


[![](https://hddtvn.com/wp-content/uploads/2021/01/ZjIB3JH.png)](https://hddtvn.com/wp-content/uploads/2021/01/ZjIB3JH.png)


Ta có công thức để đếm các bạn đã có điểm: **=COUNT(C2:C9)**


![](https://hddtvn.com/wp-content/uploads/2021/01/1k3Mcas.png)


Ta có thể thấy số bạn có điểm thi là 6 bạn. Trung vắng mặt không có điểm nên không được đếm.


Ví dụ 2: Ta có 2 bảng điểm thi toán. Cần đếm tổng số bạn đã có điểm ở cả 2 bảng.


![](https://hddtvn.com/wp-content/uploads/2021/01/SbME6L7.png)


Lúc này ta có công thức đếm tổng số bạn đã có điểm ở cả 2 bảng: **=COUNT(C2:C9;G2:G9)**


![](https://hddtvn.com/wp-content/uploads/2021/01/SeqHrqp.png)


### 2. Sử dụng hàm COUNTA để đếm số ô có ký tự


Hàm COUNTA đếm các ô chứa bất kỳ kiểu thông tin nào, gồm cả giá trị lỗi và văn bản trống (“”). Ví dụ, nếu phạm vi chứa một công thức trả về chuỗi trống, thì hàm COUNTA sẽ đếm giá trị đó. Hàm COUNTA không đếm các ô trống.


Cú pháp: **=COUNTA(value1;[value2];…)**


*Trong đó:*




* **value1**: Bắt buộc, là các ô tham chiếu hoặc các mảng trong đó bạn muốn đếm ô.

* **value2**,…:Tùy chọn, tối đa 255 mục, là các ô tham chiếu hoặc các mảng trong đó bạn muốn đếm ô.



Cũng với ví dụ trên, nếu dùng hàm COUNTA ta sẽ được kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/wP0cQ4o.png)


Ta thấy kết quả nhận được sẽ là 7 chứ không phải 6 như sử dụng hàm COUNT. Lúc này ô điểm thi của Trung cũng được đếm vì ô chứa ký tư.


### 3. Sử dụng hàm COUNTBLANK để đếm ô trống


Hàm COUNTBLANK được dùng để đếm các ô không chứa bất kỳ ký tự nào của mảng.


Cú pháp hàm: **=COUNTBLANK(range)**


Trong đó: **Range** là mảng mà bạn muốn đếm các ô trống.


*Lưu ý*: Số 0 không phải là blank (ô rỗng).


Ví dụ: Cũng với ví dụ trên, ta cần đếm số bạn chưa có điểm


![](https://hddtvn.com/wp-content/uploads/2021/01/edsQFKQ.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm COUNT, COUNTA, COUNTBLANK để đếm số ô trong Excel. Chúc các bạn thành công!*


moreKhi làm việc trên Excel, nhiều lúc bạn sẽ cần đếm các ô có giá…

