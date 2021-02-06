Cách dùng hàm NETWORKDAYS để tính số ngày làm việc của nhân viên
================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-networkdays-de-tinh-so-ngay-lam-viec-cua-nhan-vien/)Hàm NETWORKDAYS thường được sử dụng trong việc tính lương cho nhân viên một tổ chức, doanh nghiệp. Hàm NETWORKDAYS sẽ trả về tổng số ngày làm việc trừ ngày nghỉ Lễ và 2 ngày cuối tuần. Nhưng bạn vẫn có thể điều chỉnh lại công thức để đưa vào số ngày nghỉ tương ứng …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm NETWORKDAYS thường được sử dụng trong việc tính lương cho nhân viên một tổ chức, doanh nghiệp.** **Hàm NETWORKDAYS sẽ trả về tổng số ngày làm việc trừ ngày nghỉ Lễ và 2 ngày cuối tuần. Nhưng bạn vẫn có thể điều chỉnh lại công thức để đưa vào số ngày nghỉ tương ứng của công ty mình.** **Hãy theo dõi bài viết sau để biết cách sử dụng hàm NETWORKDAYS trong Excel nhé.**


### 1. Cấu trúc hàm NETWORKDAYS


Cú pháp hàm: **=NETWORKDAYS(start\_date; end\_date; [holidays])**


Trong đó:




* **Start\_date**: đối số bắt buộc, là ngày bắt đầu làm việc.

* **End\_date**: đối số bắt buộc, là ngày kết thúc làm việc.

* **Holidays**: đối số tùy chọn, là những ngày nghỉ lễ.



*Lưu ý:*




* Hàm NETWORKDAYS sẽ đếm những ngày làm việc trong tuần không tính ngày cuối tuần. Tức là đếm những ngày từ Thứ 2 đến Thứ 6. Hai ngày Thứ 7 và Chủ nhật là ngày cuối tuần sẽ không được đếm.

* Nếu bất kỳ đối số nào không phải là ngày hợp lệ, hàm NETWORKDAYS sẽ trả về giá trị lỗi.



### 2. Cách sử dụng hàm NETWORKDAYS


Ví dụ ta cần đếm số ngày làm việc trong khoảng từ ngày 15/04/2019 đến ngày 15/05/2019. Trong khoảng thời gian này sẽ có 2 ngày nghỉ lễ là 30/04/2019 và 01/05/2019.


[![](https://hddtvn.com/wp-content/uploads/2021/01/YTNI2r9.png)](https://hddtvn.com/wp-content/uploads/2021/01/YTNI2r9.png)


Áp dụng cấu trúc hàm như trên, ta sẽ có công thức tính ngày làm việc giữa ngày 15/04/2019 và 15/05/2019 theo các trường hợp như sau:




* Số ngày làm việc tính từ ngày 15/04/2019 đến ngày 15/05/2019: **=NETWORKDAYS(A3;A4)**

* Số ngày làm việc tính từ ngày 15/04/2019 đến ngày 15/05/2019 với ngày 30/04/2019 là ngày nghỉ lễ không làm việc: **=NETWORKDAYS(A3;A4;A5)**

* Số ngày làm việc tính từ ngày 15/04/2019 đến ngày 15/05/2019 với 2 ngày 30/04/2019 và 01/05/2019 là ngày nghỉ lễ không làm việc: **=NETWORKDAYS(A3;A4;A5:A6)**



![](https://hddtvn.com/wp-content/uploads/2021/01/XCdmBaL.png)


*Các bạn có thể thấy, hàm NETWORKDAYS sẽ tính ra số ngày làm việc rất nhanh chóng và chính xác. Hy vọng qua bài viết này, các bạn đã có thể sử dụng thành thạo được hàm NETWORKDAYS. Chúc các bạn thành công!*


moreHàm NETWORKDAYS thường được sử dụng trong việc tính lương cho nhân viên một tổ…

