Cách sử dụng hàm FREQUENCY để tính tần suất xuất hiện của giá trị trong Excel
=============================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-frequency-de-tinh-tan-suat-xuat-hien-cua-gia-tri-trong-excel/)Trong bài viết này, hddtvn sẽ giới thiệu với bạn đọc cách sử dụng hàm FREQUENCY để tính tần suất xuất hiện của giá trị trong mảng dữ liệu Excel. 1. Cấu trúc hàm FREQUENCY Cú pháp hàm: =FREQUENCY(data\_array, bins\_array) Trong đó: data\_array: đối số bắt buộc, là mảng hoặc tham chiếu tới một tập …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn sẽ giới thiệu với bạn đọc cách sử dụng hàm FREQUENCY để tính tần suất xuất hiện của giá trị trong mảng dữ liệu Excel.**


### 1. Cấu trúc hàm FREQUENCY


Cú pháp hàm: **=FREQUENCY(data\_array, bins\_array)**


*Trong đó:*




* **data\_array**: đối số bắt buộc, là mảng hoặc tham chiếu tới một tập giá trị mà bạn muốn đếm tần suất của nó.

* **bins\_array**: đối số bắt buộc, là mảng hoặc tham chiếu tới các khoảng mà bạn muốn nhóm các giá trị trong **data\_array** vào trong đó.



*Lưu ý:*




* Nếu **data\_array** không chứa giá trị, thì hàm FREQUENCY trả về mảng các số không.

* Nếu **bins\_array** không chứa giá trị, thì hàm FREQUENCY trả về số thành phần trong **data\_array**.

* Hàm FREQUENCY bỏ qua các ô trống và văn bản.



### 2. Cách sử dụng hàm FREQUENCY


Ví dụ ta có bảng điểm thi toán như hình dưới. Yêu cầu tính tần suất xuất hiện điểm thi từ 1 đến 10.


![](https://hddtvn.com/wp-content/uploads/2021/01/sJsTCd5.png)


Đầu tiên ta cần tạo cột dữ liệu điểm từ 1 đến 10.


![](https://hddtvn.com/wp-content/uploads/2021/01/Asz88SL.png)


Tiếp theo, áp dụng cấu trúc hàm như trên ta có công thức tìm tần suất xuất hiện điểm 1 như sau:


**=FREQUENCY(D2:D11;E2:E11)**


Ta có thể thấy không ai có điểm 1 môn toán nên giá trị trả về sẽ là 0.


![](https://hddtvn.com/wp-content/uploads/2021/01/qjpNJrR.png)


Để tìm tần suất cho các điểm còn lại, các bạn bôi đen các ô từ F2 đến F11 rồi nhấn chuột vào công thức của ô F2. Sau đó sử dụng tổ hợp phím **Ctrl + Shift + Enter** để áp dụng công thức mảng cho tất cả ô còn lại. Kết quả ta sẽ thu được tần suất xuất hiện của các điểm như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/sbS2KaQ.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm FREQUENCY để tính tần suất xuất hiện của giá trị trong mảng của Excel. Hy vọng qua bài viết này, các bạn đã nắm vững các sử dụng hàm này. Chúc các bạn thành công!*


moreTrong bài viết này, Ketoan.vn sẽ giới thiệu với bạn đọc cách sử dụng hàm…

