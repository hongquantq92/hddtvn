Cách dùng hàm REPLACE để thay thế một phần của chuỗi văn bản
================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-replace-de-thay-the-mot-phan-cu%cc%89a-chuo%cc%83i-van-ba%cc%89n/)Hàm REPLACE được sử dụng trong Excel để thay thế một phần của chuỗi văn bản bằng một chuỗi văn bản khác. Ví dụ bạn cần thay thế một đoạn văn bản cũ, chèn thêm văn bản mới vào giữa hay thay thế vào vị trí xuất hiện của văn bản thì hàm REPLACE sẽ …
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm REPLACE được sử dụng trong Excel để thay thế một phần của chuỗi văn bản bằng một chuỗi văn bản khác.** **Ví dụ bạn cần thay thế một đoạn văn bản cũ, chèn thêm văn bản mới vào giữa hay thay thế vào vị trí xuất hiện của văn bản thì hàm REPLACE sẽ giúp bạn làm điều đó.** **Bài viết sau đây sẽ hướng dẫn các bạn cách sử dụng hàm REPLACE trong Excel.**


### 1. Cấu trúc hàm REPLACE


Cú pháp hàm: **=REPLACE(old\_text; start\_num; num\_chars; new\_text)**


*Trong đó:*




* **Old\_text**: đối số bắt buộc, là văn bản hoặc tham chiếu đến một văn bản mà bạn muốn thay thế ký tự trong đó.

* **Start\_num**: đối số bắt buộc, là vị trí ký tự đầu tiên trong **old\_text** mà bạn muốn thay thế bằng văn bản mới.

* **Num\_chars**: đối số bắt buộc, là số lượng ký tự trong **old\_text** mà bạn muốn thay thế.

* **New\_text**: đối số bắt buộc, là văn bản sẽ thay thế cho **old\_text**. Nếu viết thẳng **new\_text** vào hàm thì cần để trong dấu nháy kép””.



*Lưu ý:*




* Nếu đối số **start\_num** hoặc **num\_chars** là số âm hoặc không phải số, hàm Replace trả kết quả lỗi #VALUE!.

* Hàm REPLACE luôn trả kết quả là chuỗi văn bản chứ không phải số. Bởi vì đó là một giá trị văn bản nên không thể sử dụng trong các phép tính trừ khi bạn chuyển nó trở lại số, ví dụ nhân với 1.



### 2. Cách sử dụng hàm REPLACE


Ví dụ ta cần thay thế các văn bản theo các yêu cầu như sau:




* Thay thế Bia loa thành Bia lon

* Thay thế Nước khoáng thành Nước tinh khiết

* Thay thế Năm 2019 thành Năm 2020

* Thay thế 0981421689 thành 0981.421.689



[![](https://hddtvn.com/wp-content/uploads/2021/01/EPhvdtb.png)](https://hddtvn.com/wp-content/uploads/2021/01/EPhvdtb.png)


Áp dụng cấu trúc hàm như trên, ta có công thức để thay thế Bia loa thành Bia lon như sau:


**=REPLACE(A3;7;1;”n”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/27UhFB9.png)


Ta có công thức để thay thế Nước khoáng thành Nước tinh khiết như sau:


**=REPLACE(A4;6;6;”tinh khiết”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/ZUmIDjk.png)


Ta có công thức để thay thế Năm 2019 thành Năm 2020 như sau:


**=REPLACE(A5;7;2;20)**


![](https://hddtvn.com/wp-content/uploads/2021/01/8n2rW4n.png)


Để thay thế 0981421689 thành 0981.421.689, ta có công thức dùng 2 hàm REPLACE liên tiếp nhau như sau:


**=REPLACE(REPLACE(A6;5;0;”.”);9;0;”.”)**


![](https://hddtvn.com/wp-content/uploads/2021/01/tW8cj8i.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm REPLACE trong Excel. Chúc các bạn thành công!*


moreHàm REPLACE được sử dụng trong Excel để thay thế một phần của chuỗi văn…

