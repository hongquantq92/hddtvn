Hướng dẫn cách sử dụng hàm FIND trong Excel
===========================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-cach-su-dung-ham-find-trong-excel/)Hàm FIND trong Excel là hàm tìm ký tự trong một chuỗi văn bản, cũng có mục đích như với các hàm tách chuỗi như hàm MID trong Excel hay hàm LEFT hay hàm RIGHT nhưng về cơ bản thì khác nhau hoàn toàn. Hàm FIND sẽ tự động tìm ra vị trí của một …
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Hàm FIND trong Excel là hàm tìm ký tự trong một chuỗi văn bản, cũng có mục đích như với các hàm tách chuỗi như hàm MID trong Excel hay hàm LEFT hay hàm RIGHT nhưng về cơ bản thì khác nhau hoàn toàn. Hàm FIND sẽ tự động tìm ra vị trí của một ký tự trong chuỗi văn bản còn hàm tách chuỗi trên thì người dùng cần xác định trước số ký tự và vị trí cần tách. Bài viết dưới đây sẽ hướng dẫn bạn đọc cách sử dụng hàm FIND trong Excel cũng như một số ví dụ của hàm.**


![](https://hddtvn.com/wp-content/uploads/2021/01/find.png)


### 1. Cấu trúc của hàm FIND


Hàm FIND trong Excel được sử dụng để trả lại vị trí của một ký tự hay chuỗi phụ trong một chuỗi văn bản.


Hàm FIND luôn luôn đếm mỗi ký tự là 1, cho dù đó là byte đơn hay byte kép, bất kể thiết đặt ngôn ngữ mặc định là gì.


Cú pháp hàm: **=FIND(find\_text; within\_text; [start\_num])**


*Trong đó:*




* **Find\_text**: đối số bắt buộc, ký tự hoặc chuỗi phụ bạn muốn tìm.

* **Within\_text**: đối số bắt buộc, chuỗi văn bản được tìm kiếm. Thông thường nó được xem như một ô tham chiếu, nhưng bạn cũng có thể gõ chuỗi trực tiếp vào công thức.

* **Start\_num**: đối số tùy chọn xác định vị trí của ký tự mà bạn bắt đầu tìm kiếm. Nếu không nhập, Excel sẽ tìm kiếm bắt đầu từ ký tự thứ nhất của chuỗi **Within\_text**.



*Lưu ý:*




* Hàm **FIND** phân biệt chữ hoa, chữ thường và không cho phép sử dụng ký tự đại diện

* Hàm **FIND** không cho phép sử dụng ký tự thay thế .

* Nếu đối số **find\_text** chứa nhiều ký tự, hàm **FIND** sẽ trả về vị trí của ký tự đầu tiên . Ví dụ, công thức **FIND(“ap”;”happy”)** trả về 2 vì **“a”** là ký tự thứ hai trong từ **“happy”**.

* Nếu trong phần **within\_text** chứa nhiều lần xuất hiện của tệp tin **find\_text**, lần xuất hiện đầu tiên sẽ được trả về. Ví dụ, **FIND(“l”;”hello”)** trả về 3, là vị trí của chữ **“l”** đầu tiên trong từ **“hello”**.

* Nếu **find\_text** là một chuỗi trống, hàm **FIND**  sẽ trả về ký tự đầu tiên trong chuỗi tìm kiếm.

* Hàm **FIND** sẽ trả về **#VALUE!** nếu xảy ra bất kỳ trường hợp nào sau đây:  

**Find\_text** không tồn tại trong **within\_text**.  

**Start\_num** chứa nhiều ký tự hơn **within\_text**.  

**Start\_num** là **0** (không) hoặc một số âm.



### 2. Cách sử dụng hàm FIND trong Excel


Ví dụ ta có bảng dữ liệu sau:


[![](https://hddtvn.com/wp-content/uploads/2021/01/361t5ZQ.png)](https://hddtvn.com/wp-content/uploads/2021/01/361t5ZQ.png)


#### a. Tìm vị trí của Chữ T trong ô B2


Ta có công thức **=FIND(“T”;B2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/hJtwKF5.png)


#### b. Tìm vị trí của số 0 đầu tiên trong ô B2


Ta có công thức **=FIND(0;B2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/jHQYZCp.png)


#### c. Tìm vị trí của số 0 thứ nhất bắt đầu từ ký tự thứ 5 trong ô B2


Ta có công thức: **=FIND(0;B2;5)**


![](https://hddtvn.com/wp-content/uploads/2021/01/DIjR0fV.png)


#### d. Tìm vị trí của chuỗi “hộc” trong ô C2


Ta có công thức: **=FIND(“hộc”;C2)**


![](https://hddtvn.com/wp-content/uploads/2021/01/cj2Bgur.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng hàm FIND để tìm vị trí của ký tự trong Excel. Chúc các bạn thành công!*


moreHàm FIND trong Excel là hàm tìm ký tự trong một chuỗi văn bản, cũng…

