Cách phân biệt/hướng dẫn dùng hàm COLUMN và COLUMNS trong Excel
===============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-phan-biet-huong-dan-dung-ham-column-va-columns-trong-excel/)Bạn muốn biết tham chiếu của mình thuộc cột nào hoặc tham chiếu của mình có bao nhiêu cột? Bài viết sau đây sẽ giới thiệu cho bạn cách sơ lược nhất về hàm COLUMN trả về vị trí cột của tham chiếu và hàm COLUMNS trả về số lượng cột của tham chiếu. Cả …
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn muốn biết tham chiếu của mình thuộc cột nào hoặc tham chiếu của mình có bao nhiêu cột? Bài viết sau đây sẽ giới thiệu cho bạn cách sơ lược nhất về hàm COLUMN trả về vị trí cột của tham chiếu và hàm COLUMNS trả về số lượng cột của tham chiếu. Cả hai hàm này đều có thể sử dụng độc lập hoặc kết hợp với các hàm và công thức khác.**


### 1. Hàm COLUMN


#### a. Cấu trúc hàm COLUMN


Hàm COLUMN trả về vị trí cột của tham chiếu ô đã cho hoặc cột đầu tiên của mảng tham chiếu.


Cú pháp hàm: **=COLUMN([reference])**


*Trong đó:* **reference** là đối số tùy chọn. Ô hoặc phạm vi ô mà bạn muốn trả về số cột.


*Lưu ý:*




* Nếu bỏ qua đối số tham chiếu, COLUMN sẽ trả về vị trí cột tại ô nhập công thức vào

* Nếu tham chiếu là một ô, COLUMN sẽ trả về vị trí cột của ô đó

* Trong trường hợp tham chiếu là một vùng ngang, COLUMN sẽ trả về vị trí cột của ô đầu tiên trong vùng đó

* Đối số tham chiếu không thể tham chiếu đến nhiều vùng.



#### b. Cách sử dụng hàm COLUMN


Ví dụ ta có bảng dữ liệu như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/VeUdiC3.png)


Nếu ta đặt công thức **=COLUMN(B3)**. Ta sẽ thu được kết quả là vị trí cột của ô B3


![](https://hddtvn.com/wp-content/uploads/2021/01/bZw12dw.png)


Còn nếu ta đặt công thức là **=COLUMN(C3:D3)**. Ta sẽ thu được kết quả là vị trí cột của ô đầu tiên của mảng tham chiếu. Tức là vị trí cột của ô C3.


![](https://hddtvn.com/wp-content/uploads/2021/01/11Dpaz4.png)


### 2. Hàm COLUMNS


#### a. Cấu trúc hàm COLUMNS


Hàm COLUMNS trả về số cột trong một mảng hoặc tham chiếu.


Cú pháp hàm: **=COLUMNS(array)**


*Trong đó:* **Array** là đối số bắt buộc. Mảng hoặc công thức mảng hoặc tham chiếu đến phạm vi ô mà bạn muốn số cột cho nó.


#### b. Cách sử dụng hàm COLUMNS


Cũng với ví dụ trên, nếu ta đặt công thức **=COLUMNS(C3:D3)**. Thì ta sẽ thu được kết quả là số cột của mảng tham chiếu chứ không phải vị trí cột của ô đầu tiên của mảng tham chiếu như hàm COLUMN.


![](https://hddtvn.com/wp-content/uploads/2021/01/2iERfge.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng và phân biệt sự khác nhau giữa hàm COLUMN và hàm COLUMNS. Chúc các bạn thành công!*


moreBạn muốn biết tham chiếu của mình thuộc cột nào hoặc tham chiếu của mình…

