Cách sử dụng hàm WEEKDAY để đổi ngày sang thứ trong Excel
=========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-weekday-de-doi-ngay-sang-thu-trong-excel/)Trong bài viết này, hddtvn sẽ giới thiệu tới bạn đọc cách sử dụng hàm WEEKDAY để đổi ngày sang thứ trong Excel một cách đơn giản. 1. Cấu trúc hàm WEEKDAY Cú pháp hàm: =WEEKDAY(serial\_number; [return\_type]) Trong đó: Serial\_number: đối số bắt buộc, là ngày mà bạn muốn tìm thứ trong tuần. Return\_type: đối số …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong bài viết này, hddtvn sẽ giới thiệu tới bạn đọc cách sử dụng hàm WEEKDAY để đổi ngày sang thứ trong Excel một cách đơn giản.**


### 1. Cấu trúc hàm WEEKDAY


Cú pháp hàm: **=WEEKDAY(serial\_number; [return\_type])**


*Trong đó:*




* **Serial\_number**: đối số bắt buộc, là ngày mà bạn muốn tìm thứ trong tuần.

* **Return\_type**: đối số tùy chọn, là một số xác định kiểu giá trị trả về.



*Lưu ý:*




* Hàm sẽ trả về giá trị lỗi #NUM! nếu **serial\_number** có giá trị vượt quá giá trị ngày tháng

* Hàm sẽ trả về giá trị lỗi #VALUE! nếu **serial\_number** không phải dạng số hoặc ngày tháng

* Nếu **return\_type** bằng 1 hoặc bỏ qua thì giá trị trả về là các số từ 1 (chủ nhật) đến 7 (thứ bảy)

* Nếu **return\_type** bằng 2 thì giá trị trả về là các số từ 1 (thứ hai) đến 7 (chủ nhật)

* Nếu **return\_type** bằng 3 thì giá trị trả về là các số từ 0 (thứ hai) đến 6 (chủ nhật)

* Nếu **return\_type** bằng 11 thì giá trị trả về là các số từ 1 (thứ hai) đến 7 (chủ nhật)

* Nếu **return\_type** bằng 12 thì giá trị trả về là các số từ 1 (thứ ba) đến 7 (thứ hai)

* Nếu **return\_type** bằng 13 thì giá trị trả về là các số từ 1 (thứ tư) đến 7 (thứ ba)

* Nếu **return\_type** bằng 14 thì giá trị trả về là các số từ 1 (thứ năm) đến 7 (thứ tư)

* Nếu **return\_type** bằng 15 thì giá trị trả về là các số từ 1 (thứ sáu) đến 7 (thứ năm)

* Nếu **return\_type** bằng 16 thì giá trị trả về là các số từ 1 (thứ bảy) đến 7 (thứ sáu)

* Nếu **return\_type** bằng 17 thì giá trị trả về là các số từ 1 (chủ nhật) đến 7 (thứ bảy)



### 2. Cách sử dụng hàm WEEKDAY để đổi ngày sang thứ


Ví dụ ta cần đổi các ngày trong hình dưới sang thứ.


![](https://hddtvn.com/wp-content/uploads/2021/01/hylqa5d.png)


Áp dụng cấu trúc hàm như trên ta có công thức như sau:


**=WEEKDAY(A2)**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả như hình dưới. Vì trong công thức ta bỏ qua đối số **return\_type** nên kết quả trả về 1 sẽ tương ứng với chủ nhật, 2 tương ứng với thứ 2… và 7 tương ứng với thứ 7.


![](https://hddtvn.com/wp-content/uploads/2021/01/fLRvPzG.png)


Tiếp theo, chúng ta chuyển thứ sang dạng chữ bằng công thức sau:


**=IF(B2=1;”Chủ nhật”;”Thứ “&B2)**


Sao chép công thức cho các ô còn lại ta sẽ thu được kết quả như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/2UG6YxL.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách đổi ngày sang thứ trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreTrong bài viết này, Ketoan.vn sẽ giới thiệu tới bạn đọc cách sử dụng hàm…

