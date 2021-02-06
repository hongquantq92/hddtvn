Hướng dẫn tính số ngày công, ngày nghỉ, ngày làm việc trong Excel
=================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-tinh-so-ngay-cong-ngay-nghi-ngay-lam-viec-trong-excel/)Ngày công, ngày nghỉ, ngày làm việc là căn cứ để tính tiền lương, thưởng cho người lao động. Đây cũng là công việc cơ bản và thường xuyên phải làm của kế toán trong doanh nghiệp. Tuy nhiên làm sao để tính số ngày công, ngày nghỉ, ngày làm việc chính xác nhất mới …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Ngày công, ngày nghỉ, ngày làm việc là căn cứ để tính tiền lương, thưởng cho người lao động. Đây cũng là công việc cơ bản và thường xuyên phải làm của kế toán trong doanh nghiệp. Tuy nhiên làm sao để tính số ngày công, ngày nghỉ, ngày làm việc chính xác nhất mới là vấn đề được kế toán quan tâm. Bài viết sẽ hướng dẫn các bạn cách sử dụng hàm Excel để tính số ngày công, ngày nghỉ, ngày làm việc nhanh và chính xác nhất nhé.**


![](https://hddtvn.com/wp-content/uploads/2021/01/Cach-tinh-so-ngay-cong-ngay-nghi-ngay-lam-viec-trong-Excel-1.png)


### 1. Cấu trúc hàm WORKDAY.INTL


Cú pháp hàm: **=WORKDAY.INTL(start\_date, days, [weekend], [holidays])**


*Trong đó:*




* **Start\_date**: đối số bắt buộc, là ngày bắt đầu.

* **Days**: đối số bắt buộc, là số ngày làm việc trước hoặc sau ngày bắt đầu. Giá trị dương cho kết quả là một ngày trong tương lai; giá trị âm cho kết quả là một ngày trong quá khứ; giá trị 0 cho kết quả là ngày bắt đầu.

* **Weekend**: đối số tùy chọn, là những ngày nào trong tuần là ngày cuối tuần và không được coi là ngày làm việc. Weekend là một số của ngày cuối tuần hoặc một chuỗi chỉ rõ khi nào thì diễn ra ngày cuối tuần.

* **Holidays**: đối số tùy chọn, là một tập hợp tùy chọn gồm một hoặc nhiều ngày cần được trừ khỏi lịch ngày làm việc.



*Các giá trị của Weekend:*




* Weekend = 1 => Ngày cuối tuần là thứ 7, chủ nhật.

* Weekend = 2 => Ngày cuối tuần là chủ nhật, thứ hai.

* Weekend = 3 => Ngày cuối tuần là thứ hai, thứ ba.

* Weekend = 4 => Ngày cuối tuần là thứ thứ ba, thứ tư.

* Weekend = 5 => Ngày cuối tuần là thứ thứ tư, thứ năm.

* Weekend = 6 => Ngày cuối tuần là thứ năm, thứ sáu.

* Weekend = 7 => Ngày cuối tuần là thứ sáu, thứ 7.

* Weekend = 11 => Ngày cuối tuần chỉ có chủ nhật.

* Weekend = 12 => Ngày cuối tuần chỉ có thứ hai.

* Weekend = 13 => Ngày cuối tuần chỉ có thứ ba.

* Weekend = 14 => Ngày cuối tuần chỉ có thứ tư.

* Weekend = 15 => Ngày cuối tuần chỉ có thứ năm.

* Weekend = 16 => Ngày cuối tuần chỉ có thứ sáu.

* Weekend = 17 => Ngày cuối tuần chỉ có thứ 7.



### 2. Cách sử dụng hàm WORKDAY.INTL trong Excel


Ví dụ ta cần tìm ngày làm việc thứ 60 tính từ ngày bắt đầu là 01/08/2019. Trong đó nghỉ cuối tuần thứ 7 và chủ nhật, cũng như được nghỉ lễ quốc khánh 02/09/2019.


Áp dụng cấu trúc hàm bên trên ta sẽ có công thức tính ra ngày làm việc thứ 60 như sau:


**=WORKDAY.INTL(A2;B2;C2;D2)**


Trong đó đối số cuối tuần có giá trị là 1 do ta đặt cuối tuần là thứ 7 và chủ nhật.


![](https://hddtvn.com/wp-content/uploads/2021/01/20-4.png)


Tương tự với ví dụ trên nhưng nếu ta đặt cuối tuần chỉ là chủ nhật thì đối số cuối tuần sẽ có giá trị là 11. Khi đó ngày làm việc thứ 60 sẽ ngày 11/10/2019.


![](https://hddtvn.com/wp-content/uploads/2021/01/21-5.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách tính số ngày công, ngày nghỉ, ngày làm việc trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreNgày công, ngày nghỉ, ngày làm việc là căn cứ để tính tiền lương, thưởng…

