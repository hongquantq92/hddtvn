Cách dùng Hàm ROW (và ROWS) để trả về vị trí (số lượng) hàng của tham chiếu
===========================================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-dung-ham-row-va-rows-de-tra-ve-vi-tri-so-luong-hang-cua-tham-chieu/)Bạn muốn biết tham chiếu của mình thuộc hàng nào hoặc tham chiếu của mình có bao nhiêu hàng? Hãy sử dụng tới bô đôi hàm tham chiếu/tra cứu là ROW và ROWS nhé. Hàm ROW sẽ trả về vị trí hàng của tham chiếu và hàm ROWS trả về số lượng hàng của tham …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Bạn muốn biết tham chiếu của mình thuộc hàng nào hoặc tham chiếu của mình có bao nhiêu hàng? Hãy sử dụng tới bô đôi hàm tham chiếu/tra cứu là ROW và ROWS nhé. Hàm ROW sẽ trả về vị trí hàng của tham chiếu và hàm ROWS trả về số lượng hàng của tham chiếu. Là các hàm trang tính, hàm ROW và ROWS có thể được nhập như một phần của công thức trong một ô của trang tính.**


### 1. Hàm ROW


#### a. Cấu trúc hàm ROW


Hàm ROW trả về vị trí hàng của tham chiếu ô đã cho hoặc hàng đầu tiên của mảng tham chiếu.


Cú pháp hàm: **=ROW([reference])**


Trong đó: **reference** là đối số tùy chọn. Ô hoặc phạm vi ô mà bạn muốn trả về số hàng.


*Lưu ý:*




* Nếu bỏ qua đối số tham chiếu, ROW sẽ trả về vị trí hàng tại ô nhập công thức vào

* Nếu tham chiếu là một ô, ROW sẽ trả về vị trí cột của ô đó

* Trong trường hợp tham chiếu là một vùng dọc, ROW sẽ trả về vị trí hàng của ô đầu tiên trong vùng đó

* Đối số tham chiếu không thể tham chiếu đến nhiều vùng.



#### b. Cách sử dụng hàm ROW


Ví dụ ta có bảng dữ liệu như sau:


![](https://hddtvn.com/wp-content/uploads/2021/01/lu7Xzdz.png)


Nếu ta đặt công thức **=ROW(B3)**. Ta sẽ thu được kết quả là vị trí hàng của ô B3 là 3


![](https://hddtvn.com/wp-content/uploads/2021/01/uxoLfk6.png)


Còn nếu ta đặt công thức là **=ROW(B4:B10)**. Ta sẽ thu được kết quả là vị trí hàng của ô đầu tiên của mảng tham chiếu. Tức là vị trí cột của ô B4 là 4


![](https://hddtvn.com/wp-content/uploads/2021/01/g7zoNeT.png)


### 2. Hàm ROWS


#### a. Cấu trúc hàm ROWS


Hàm ROWS trả về số lượng hàng trong một mảng hoặc tham chiếu.


Cú pháp hàm: **=ROWS(array)**


Trong đó: **Array** đối số bắt buộc, là mảng hoặc công thức mảng hoặc tham chiếu đến phạm vi ô mà bạn muốn lấy số hàng cho nó.


#### b. Cách sử dụng hàm ROWS


Cũng với ví dụ trên, nếu ta đặt công thức **=ROWS(B4:B10)**. Thì ta sẽ thu được kết quả là số hàng của mảng tham chiếu là 7 chứ không phải vị trí hàng của ô đầu tiên của mảng tham chiếu như hàm ROW.


![](https://hddtvn.com/wp-content/uploads/2021/01/jlMomf3.png)


*Như vậy, bài viết trên đã hướng dẫn các bạn cách sử dụng và phân biệt sự khác nhau giữa hàm ROW và hàm ROWS. Chúc các bạn thành công!*


moreBạn muốn biết tham chiếu của mình thuộc hàng nào hoặc tham chiếu của mình…

