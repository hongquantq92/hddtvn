Cách dùng hàm DB để tính khấu hao TSCĐ theo số dư giảm dần cố định
==================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-db-de-tinh-khau-hao-tscd-theo-so-du-giam-dan-co-dinh/)Trong quá trình sử dụng và đầu tư các thiết bị kĩ thuật các bạn phải xem xét mức chi phí và tính phí hao mòn tài sản khi đi vào sử dụng. Lúc này, hàm DB sẽ giúp bạn nhận thấy giá trị khấu hao tài sản giảm dần qua các năm, từ đó …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Trong quá trình sử dụng và đầu tư các thiết bị kĩ thuật các bạn phải xem xét mức chi phí và tính phí hao mòn tài sản khi đi vào sử dụng.** **Lúc này, hàm DB sẽ giúp bạn nhận thấy giá trị khấu hao tài sản giảm dần qua các năm, từ đó sẽ giúp các bạn có kế hoạch sử dụng tài sản một cách hợp lý. Bài viết sẽ hướng dẫn các bạn cách sử dụng hàm DB để tính khấu hao tài sản cố định theo phương pháp số dư giảm dần cố định.**


### 1. Cấu trúc hàm DB


Cú pháp hàm: **=DB(cost; salvage; life; period; [month])**


*Trong đó:*




* **Cost**: đối số bắt buộc., là nguyên giá của tài sản.

* **Salvage**: đối số bắt buộc, là giá trị còn lại của tài sản (hay còn được gọi là giá trị thu hồi của tài sản).

* **Life**: đối số bắt buộc, là số kỳ tính khấu hao của tài sản (hay còn được gọi là tuổi thọ hữu ích của tài sản).

* **Period**: đối số bắt buộc, là số kỳ mà bạn muốn tính khấu hao.

* **Month**: đối số tùy chọn, là số tháng trong năm đầu tiên tính khấu hao.



*Lưu ý:*




* **Period** phải cùng đơn vị tính với **Life**

* Nếu bỏ qua đối số **month**, nó được mặc định là 12 tức là ngày bắt đầu tính khấu hao tài sản là ngày đầu tiên của năm.

* Hàm **DB** công thức sau đây để tính toán khấu hao trong một kỳ:

* (**cost** – tổng số khấu hao từ các kỳ trước) * (1 – ((**salvage** / **cost**) ^ (1 / **life**)))

* Đối với kỳ đầu tiên, hàm **DB** dùng công thức sau: **cost** * **rate** * **month** / 12

* Đối với kỳ cuối cùng, hàm DB dùng công thức sau: ((**cost** – tổng khấu hao từ các kỳ trước) * **rate** * (12 – **month**)) / 12

* Hàm **DB** sẽ trả về kết quả dạng tiền tệ, nếu không muốn các bạn có thể định dạng lại ô.



### 2. Cách sử dụng hàm DB


Ví dụ ta có danh sách tài sản cần tính khấu hao theo phương pháp số dư giảm dần cố định như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/5aJhrJq.png)


Áp dụng cấu trúc hàm DB như trên, ta có công thức tính khấu hao Cổng chào trong năm 1 như sau:


**=DB(C2;D2;E2;1)**


![](https://hddtvn.com/wp-content/uploads/2021/01/tjue7Xh.png)


Tương tự năm 1, ta có công thức tính khấu hao theo phương pháp số dư giảm dần cố định cho các năm tiếp theo như sau:




* Năm 2: **=DB(C2;D2;E2;2)**

* Năm 3: **=DB(C2;D2;E2;3)**

* Năm 4: **=DB(C2;D2;E2;4)**

* Năm 5: **=DB(C2;D2;E2;5)**



Sao chép công thức tính cho các tài sản còn lại ta thu được kết quả:


![](https://hddtvn.com/wp-content/uploads/2021/01/i3Veu84.png)


Nếu không muốn sử dụng đơn vị tiền tệ, các bạn có thể định dạng lại số sang dạng thập phân bằng cách nhấn vào biểu tượng dấu phẩy tại mục **Number** trong thẻ **Home**.


![](https://hddtvn.com/wp-content/uploads/2021/01/S3Clhz4.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm DB để tính khấu hao TSCĐ theo phương pháp số dư giảm dần cố định trong Excel. Chúc các bạn thành công!*


moreTrong quá trình sử dụng và đầu tư các thiết bị kĩ thuật các bạn…

