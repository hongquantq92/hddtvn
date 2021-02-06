Vì sao Excel chỉ hiển thị công thức, không trả về kết quả?
==========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/vi-sao-excel-chi-hien-thi-cong-thuc-khong-tra-ve-ket-qua/)Đã bao giờ bạn gặp tình trạng nhập công thức vào Excel mà Excel không trả về kết quả, chỉ hiển thị nội dung công thức? Cùng hddtvn tìm hiểu nguyên nhân tại bài viết dưới đây để khắc phục nhé. Do bật tính năng Show Formulas Khi bạn nhập công thức mà toàn bộ …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Đã bao giờ bạn gặp tình trạng nhập công thức vào Excel mà Excel không trả về kết quả, chỉ hiển thị nội dung công thức? Cùng hddtvn tìm hiểu nguyên nhân tại bài viết dưới đây để khắc phục nhé.**


![Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?](https://hddtvn.com/wp-content/uploads/2021/01/nsAT7v7.png "Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?")


Do bật tính năng **Show Formulas**
----------------------------------


Khi bạn nhập công thức mà toàn bộ đều chỉ hiển thị công thức chứ không hiển thị kết quả có nghĩa là bạn đang bật tính năng **Show Formulas** trong thẻ **Formulas** trên thanh công cụ. Tính năng này sẽ hiển thị toàn bộ công thức có trong trang tính, đồng nghĩa với nó là độ rộng cột sẽ được mở rộng.


![Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?](https://hddtvn.com/wp-content/uploads/2021/01/6eMMQT5.png "Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?")


![](https://hddtvn.com/wp-content/uploads/2021/01/nsAT7v7.png)


Do công thức có định dạng TEXT
------------------------------


Nếu công thức của bạn có định dạng Text, Excel sẽ không trả về kết quả mà chỉ hiển thị công thức. Một cách để xác định công thức của bạn có đang ở định dạng Text hay không là sử dụng hàm ISTEXT. Nếu hàm của bạn có định dạng Text thì kết quả trả về sẽ là TRUE.


Ví dụ như hình minh họa dưới đây:


![Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?](https://hddtvn.com/wp-content/uploads/2021/01/Q01IpHe.png "Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?")


Nhưng tại sao công thức của bạn lại có định dạng Text. Hãy cùng điểm qua 1 số lỗi cơ bản dẫn đến tình trạng này dưới đây:


#### Thiếu dấu “=” trước công thức


Khi bạn quên nhập dấu bằng “=” trước công thức thì công thức sẽ trở thành định dạng Text và không hiển thị kết quả.


Ví dụ dưới đây thiếu dấu bằng trước hàm ROUND nên không hiển thị kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/zXlXZSS.png)


#### Dấu cách trước dấu “=”


Trong một sô trường hợp, bạn ấn nhầm dấu cách trước khi nhập công thức nên Excel sẽ không hiển thị kết quả.


![Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?](https://hddtvn.com/wp-content/uploads/2021/01/ypZg7hP.png "Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?")


#### Đặt công thức trong dấu nháy kép ” “


Khi đặt công thức vào trong dấu kép như ví dụ “=ROUND(E4,3)” dưới đây thì Excel cũng không hiển thị kết quả.


![](https://hddtvn.com/wp-content/uploads/2021/01/yrnkPom.png)


### Đặt dấu nháy đơn ở đầu công thức


Khi bạn vô tình ấn nhầm dấu nháy đơn ‘ ở trước công thức, công thức sẽ trở thành dạng Text. Do đó sẽ không hiển thị kết quả. Tình trạng này thường khó nhận biết vì khi bạn nhập công thức xong và ấn Enter thì dấu nháy sẽ biến mất như hình minh họa dưới đây.


![Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?](https://hddtvn.com/wp-content/uploads/2021/01/YafyxyR.png "Vì sao Excel không hiển thị kết quả mà chỉ hiển thị công thức?")


Trên đây là các lý do làm cho Excel không hiển thị kết quả mà chỉ hiển thị công thức. Mời bạn đọc tham khảo để tránh sai sót trong quá trình làm việc. Chúc các bạn thành công.


moreĐã bao giờ bạn gặp tình trạng nhập công thức vào Excel mà Excel không…

