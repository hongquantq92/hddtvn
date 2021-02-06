Cách sử dụng hàm WEEKDAY để tìm thứ trong tuần của một ngày
===========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-su-dung-ham-weekday-de-tim-thu-trong-tuan-cua-mot-ngay/)Hàm WEEKDAY thuộc nhóm hàm ngày giờ (Date and Time). Hàm sẽ trả về một số nguyên là biểu diễn của ngày trong tuần cho một ngày nhất định. Hàm này rất hữu ích trong phân tích. Giả sử bạn muốn xác định thời gian cần thiết để hoàn thành một dự án nhất định. …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm WEEKDAY thuộc nhóm hàm ngày giờ (Date and Time). Hàm sẽ trả về một số nguyên là biểu diễn của ngày trong tuần cho một ngày nhất định. Hàm này rất hữu ích trong phân tích. Giả sử bạn muốn xác định thời gian cần thiết để hoàn thành một dự án nhất định. Chức năng của hàm này sẽ giúp ích trong việc loại bỏ các ngày cuối tuần khỏi khung thời gian nhất định. Vì vậy, nó đặc biệt hữu ích khi lập kế hoạch và lên lịch làm việc cho các dự án kinh doanh.****Trong bài viết này, hddtvn sẽ giới thiệu tới bạn đọc cách sử dụng hàm WEEKDAY để tìm thứ trong tuần của một ngày trong Excel.**


### 1. Cấu trúc hàm WEEKDAY


Cú pháp hàm: **=WEEKDAY(serial\_number; [return\_type])**


*Trong đó:*




* **Serial\_number**: đối số bắt buộc, là ngày mà bạn muốn tìm thứ trong tuần.

* **Return\_type**: đối số tùy chọn, là một số xác định kiểu giá trị trả về.



*Lưu ý:*




* Hàm sẽ trả về giá trị lỗi #NUM! nếu **serial\_number** có giá trị vượt quá giá trị ngày tháng

* Hàm sẽ trả về giá trị lỗi #VALUE! nếu **serial\_number** không phải dạng số hoặc ngày tháng

* Nếu **return\_type** bằng 1 hoặc bỏ qua thì giá trị trả về là các số từ 1 (chủ nhật) đến 7 (thứ bảy)

* Nếu **return\_type** bằng 2 thì giá trị trả về là các số từ 1 (thứ hai) đến 7 (chủ nhật)

* Nếu **return\_type** bằng 3 thì giá trị trả về là các số từ 0 (thứ hai) đến 6 (chủ nhật)

* Nếu **return\_type** bằng 11 thì giá trị trả về là các số từ 1 (thứ hai) đến 7 (chủ nhật)

* Nếu **return\_type** bằng 12 thì giá trị trả về là các số từ 1 (thứ ba) đến 7 (thứ hai)

* Nếu **return\_type** bằng 13 thì giá trị trả về là các số từ 1 (thứ tư) đến 7 (thứ ba)

* Nếu **return\_type** bằng 14 thì giá trị trả về là các số từ 1 (thứ năm) đến 7 (thứ tư)

* Nếu **return\_type** bằng 15 thì giá trị trả về là các số từ 1 (thứ sáu) đến 7 (thứ năm)

* Nếu **return\_type** bằng 16 thì giá trị trả về là các số từ 1 (thứ bảy) đến 7 (thứ sáu)

* Nếu **return\_type** bằng 17 thì giá trị trả về là các số từ 1 (chủ nhật) đến 7 (thứ bảy)



### 2. Cách sử dụng hàm WEEKDAY


Ví dụ ta cần tìm thứ trong tuần của ngày 17/03/2019. Áp dụng cấu trúc hàm bên trên, ta có công thức cho từng trường hợp đối số **return\_type** như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/aL7L0f6.png)


Nhìn vào kết quả ta có thể thấy với đối số **return\_type** khác nhau thì kết quả trả về cũng sẽ là các con số khác nhau nhưng ý nghĩa thì vẫn là một. Trong trường hợp này thì kết quả ta thu được ngày 17/03/2019 sẽ là ngày Chủ nhật. Nhưng sẽ được thể hiện bằng các con số khác nhau theo từng trường hợp của **return\_type**.


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm WEEKDAY trong Excel. Chúc các bạn thành công!*


moreHàm WEEKDAY thuộc nhóm hàm ngày giờ (Date and Time). Hàm sẽ trả về một…

