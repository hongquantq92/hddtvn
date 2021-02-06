Mời bạn đọc tải về file Excel tính lương tăng ca làm thêm giờ
=============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/moi-ban-doc-tai-ve-file-excel-tinh-luong-tang-ca-lam-them-gio/)Tính lương tăng ca làm thêm giờ là vấn đề phức tạp do mỗi người sẽ làm theo giờ khác nhau. Hơn nữa, có rất nhiều loại tăng ca: tăng ca thêm gờ, tăng ca chủ nhật, tăng ca ngày lễ, tết phép, nghỉ bù… Trong bài viết này, hddtvn xin chia sẻ với các …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Tính lương tăng ca làm thêm giờ là vấn đề phức tạp do mỗi người sẽ làm theo giờ khác nhau. Hơn nữa, có rất nhiều loại tăng ca: tăng ca thêm gờ, tăng ca chủ nhật, tăng ca ngày lễ, tết phép, nghỉ bù… Trong bài viết này, hddtvn xin chia sẻ với các bạn mẫu file Excel tính lương tăng ca làm thêm giờ đầy đủ nhất.**


![](https://hddtvn.com/wp-content/uploads/2021/01/QOMT6tO.png)


### 1. Một số quy ước chấm công trong file




* X: Công trong giờ ngày thường 8 tiếng, nếu ít hơn 8 giờ, ghi số giờ

* P: Phép hưởng lương

* L: lễ nghỉ hưởng lương

* TC: Tăng ca chủ nhật, nếu tăng ca ít hơn 8 giờ ghi số giờ

* TCL: tăng ca lễ, nếu tăng ca ít hơn 8 giờ ghi số giờ

* NB: Nghỉ bù hưởng lương

* Chủ nhật, lễ, tăng ca nếu nghỉ bù đánh TB (không tính lương, do đó NB tính lương). Số ngày nghỉ bù tương ứng nên có sheet theo dõi riêng.

* Số giờ làm việc ghi số.



### 2. Một số quy ước tính lương từ số ngày công, giờ công (có thể khác nhau tuỳ DN)




* Ngày thường: tăng ca sau 5 giờ nhân 1.5, sau 9 giờ nhân 2

* Chủ nhật: nhân 1.5, tăng thêm giờ sau 5 giờ nhân 2, sau 9 giờ nhân 3

* Lễ: Nhân 3, tăng thêm giờ sau 5 giờ nhân 4.5

* Các hệ số nhân này ghi vào dòng 9 tại các cột tương ứng.



### 3. Quy ước khác: Tuỳ doanh nghiệp




* Phụ cấp đi lại cho 1 ngày có đi làm

* Phụ cấp tiền ăn trưa

* Phụ cấp trách nhiệm

* Phụ cấp công trình bằng tỷ lệ % so với lương chính.

* Tăng ca trên 3 giờ 1 ngày hưởng thêm phụ cấp tiền ăn tối. số giờ quy định này ghi vào ô K43 bảng chấm công



### 4. Cách sử dụng file Excel tính lương tăng ca làm thêm giờ


#### a. Bảng chấm công:




* Trong file có sử dụng công thức tính ngày trong tháng, do đó, khi đổi sang tháng mới, chỉ sửa ngày đầu tiên (ngày 1) của tháng đó.

* Tiêu đề Bảng chấm công tháng mấy, Bảng lương tháng mấy cũng là công thức

* Ngày công quy định & ngày lễ nếu có, tính bằng công thức

* Một ngày chấm công 3 cột, nếu DN nào sử dụng 2 cột thì có thể dấu bớt: Cột ký hiệu “C”, chấm công giờ hành chánh, cột T5: số giờ tăng ca từ sau 5 giờ chiều đến 9 giờ tối, cột ký hiệu T9: số giờ tăng ca sau 9 giờ

* Nếu trong tháng có ngày lễ, đánh “L” vào dòng 9, 3 cột của ngày đó.

* Trong file có sử dụng Conditional Formating, tô màu ngày chủ nhật. Nếu DN nghỉ thứ bảy, thì bổ sung.

* Công thức đang tính cho DN áp dụng tuần làm 6 ngày x 8 giờ. Nếu 1 ngày 7 giờ hoặc 6 giờ, ghi vào ô G4.



#### b. Bảng lương:




* Tên, chức vụ, bộ phận, … có thể copy từ bảng chấm công sang

* Mức lương chính & trợ cấp trách nhiệm, có thể lookup từ 1 sheet khác.

* Mức tiền phụ cấp xa nhà, phụ cấp đi lại, phụ cấp tiền ăn trưa, gõ vào các ô riêng để có thể sửa chữa, thay đổi sau này.

* Nếu mức lương đóng BHXH khác mức lương chính, chèn thêm 1 cột và sửa công thức.

* Khi chèn dòng, phải đứng tại dòng trên của dòng cộng để chèn, không được đứng ngay dòng cộng.



### Các bạn có thể tải về mẫu file Excel tính lương tăng ca **TẠI ĐÂY**.


moreTính lương tăng ca làm thêm giờ là vấn đề phức tạp do mỗi người…

