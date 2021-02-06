Hướng dẫn chi tiết cách sử dụng hàm RECEIVED trong Excel
========================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-chi-tiet-cach-su-dung-ham-received-trong-excel/)Hàm RECEIVED là một trong những hàm tài chính cơ bản được sử dụng để tính số tiền nhận được khi đáo hạn một khoản đầu tư. Hãy theo dõi bài viết sau để nắm rõ cách sử dụng hàm RECEIVED trong Excel.   1. Cấu trúc hàm RECEIVED Cú pháp hàm: =RECEIVED(settlement, maturity, investment, …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm RECEIVED là một trong những hàm tài chính cơ bản được sử dụng để tính số tiền nhận được khi đáo hạn một khoản đầu tư. Hãy theo dõi bài viết sau để nắm rõ cách sử dụng hàm RECEIVED trong Excel.**


 


### 1. Cấu trúc hàm RECEIVED


Cú pháp hàm: **=RECEIVED(settlement, maturity, investment, discount, [basis])**


*Trong đó:*




* **Settlement**: đối số bắt buộc, là ngày thanh toán chứng khoán. Ngày thanh toán chứng khoán là ngày sau ngày phát hành khi chứng khoán được bán cho người mua.

* **Maturity**: đối số bắt buộc, là ngày đáo hạn của chứng khoán. Ngày đáo hạn là ngày mà chứng khoán hết hạn.

* **Investment**: đối số bắt buộc, là số tiền đã đầu tư vào chứng khoán.

* **Discount**: đối số bắt buộc, là tỷ lệ chiết khấu của chứng khoán.

* **Basis**: đối số tùy chọn, là loại cơ sở đếm ngày sẽ dùng.



Lưu ý:




* Nếu đối số **Basis** là 0 hoặc bỏ qua thì cơ sở đếm ngày dạng US (NASD) 30/360.

* Nếu đối số **Basis** là 1 thì cơ sở đếm ngày dạng Thực tế/thực tế.

* Nếu đối số **Basis** là 2 thì cơ sở đếm ngày dạng Thực tế/360.

* Nếu đối số **Basis** là 3 thì cơ sở đếm ngày dạng Thực tế/365.

* Nếu đối số **Basis** là 4 thì cơ sở đếm ngày dạng European 30/360.

* Các đối số **settlement**, **maturity** và **basis** bị cắt phần thập phân thành số nguyên.

* Nếu **settlement** hoặc **maturity** không phải là ngày hợp lệ, hàm RECEIVED trả về giá trị lỗi #VALUE!.

* Nếu **investment** ≤ 0 hoặc nếu **discount** ≤ 0, hàm RECEIVED trả về giá trị lỗi #NUM!.

* Nếu **basis** < 0 hoặc nếu **basis** > 4, thì hàm RECEIVED trả về giá trị lỗi #NUM!.

* Nếu **settlement** ≥ **maturity**, hàm RECEIVED trả về giá trị lỗi #NUM!.



### 2. Cách sử dụng hàm RECEIVED


Ví dụ bạn mua một trái phiếu trị giá 1.000.000 phát hành ngày 23/07/2019 và đáo hạn vào ngày 23/07/2020 với chiết khấu 7%.


![](https://hddtvn.com/wp-content/uploads/2021/01/V2MFthm.png)


Lúc này, để tính số tiền nhận được sau khi trái phiếu đáo hạn ta có công thức như sau:


**=RECEIVED(“23/7/2019″;”23/7/2020”;1000000;7%;3)**


Kết quả ta sẽ thu được số tiền nhận được là 1.075.491


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm RECEIVED trong Excel. Hy vọng bài viết sẽ hữu ích với các bạn trong quá trình làm việc. Chúc các bạn thành công!*


moreHàm RECEIVED là một trong những hàm tài chính cơ bản được sử dụng để…

