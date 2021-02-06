Cách gộp 2 hay nhiều cột trong Excel mà vẫn giữ nguyên nội dung
===============================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/cach-gop-2-hay-nhieu-cot-trong-excel-ma-van-giu-nguyen-noi-dung/)Gộp 2 (hay nhiều cột) thành 1 cột duy nhất là thao tác thường xuyên được sử dụng trong Excel. Ví dụ như gộp 1 cột Họ và 1 cột Tên hay gộp cột Họ và cột Tên đệm với nhau. Điều quan trọng là nội dung được ghép từ 2 cột phải đầy đủ và …
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

**Gộp 2 (hay nhiều cột) thành 1 cột duy nhất là thao tác thường xuyên được sử dụng trong Excel. Ví dụ như gộp 1 cột Họ và 1 cột Tên hay gộp cột Họ và cột Tên đệm với nhau. Điều quan trọng là nội dung được ghép từ 2 cột phải đầy đủ và không mất dữ liệu. Trên Excel có một số tính năng gộp ô như Merge Cell, nhưng khi sử dụng sẽ bị mất 1 cột bên phải, đồng nghĩa với nội dung của cột biến mất. Vậy phải làm sao để ghép các cột trong Excel lại với nhau mà không làm mất nội dung của chúng? Bài viết dưới đây sẽ hướng dẫn chi tiết cách làm cho bạn.**


![Cách gộp 2 hay nhiều cột trong Excel mà vẫn giữ nguyên nội dung](https://hddtvn.com/wp-content/uploads/2021/01/QPiR1GJ.png "Cách gộp 2 hay nhiều cột trong Excel mà vẫn giữ nguyên nội dung")


1. Cách ghép 2 cột trong Excel
------------------------------


Bài viết này sẽ lấy ví dụ ghép từ 1 cột “Họ” và 1 cột “Tên” thành cột duy nhất.


![Cách gộp 2 hay nhiều cột trong Excel mà vẫn giữ nguyên nội dung](https://hddtvn.com/wp-content/uploads/2021/01/ovFh6Ty.png "Cách gộp 2 hay nhiều cột trong Excel mà vẫn giữ nguyên nội dung")


### Cách 1: Sử dụng “toán tử &”


Trong Excel ta có thể sử dụng một cách đơn giản nhất để gộp 2 cột với nhau là dùng “toán tử &”. Cách này cho phép nối 2 hay nhiều chỗi ký tự lại với nhau. Trong ví dụ dưới đây, chúng ta sẽ thực hiện nối ô A2 cùng 1 khoảng trắng và ô B2, sau đó hoàn chỉnh việc nhập nội dung vào ô C2.


Việc thực hiện sẽ tiến hành theo công thức sau, tại ô C2: =A2&" "&B2.


![Hướng dẫn cách ghép 2 hay nhiều cột trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/hdVOY3A.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


Khi đã nhập xong, chúng ta nhấn Enter và kết quả nhận về là ô C2 có cột Họ và tên được ghép lại với nhau từ ô A2 và B2. Để thực hiện với các ô còn lại một cách nhanh chóng, chúng ta chỉ cần copy công thức vào các ô tiếp theo để có được cột Họ và tên hoàn chỉnh.


![Hướng dẫn cách ghép 2 hay nhiều cột trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/O22XkF0.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


### Cách 2: Sử dụng hàm nối chuỗi CONCATENATE


Bên cạnh đó,chúng ta có thể ghép 2 cột thành 1 cột duy nhất trong Excel bằng cách sử dụng hàm nối chuỗi CONCATENATE.


Việc thực hiện sẽ tiến hành theo công thức sau, tại ô C2: **=CONCATENATE(A2," ",B2)** rồi nhấn Enter.


![Hướng dẫn cách ghép 2 hay nhiều cột trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/KUx1LKN.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


Kết quả nhận về là ô C2 có cột Họ và tên được ghép lại với nhau từ ô A2 và B2. Để thực hiện với các ô còn lại một cách nhanh chóng, chúng ta chỉ cần copy công thức vào các ô tiếp theo để có được cột Họ và tên hoàn chỉnh.


![](https://hddtvn.com/wp-content/uploads/2021/01/QPiR1GJ.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


2. Cách gộp nhiều cột trong Excel
---------------------------------


Bên cạnh việc ghép 2 cột trong Excel thì người dùng có thể tiến hành việc gộp nhiều cột thành một cột duy nhất bằng cách dễ dàng và gần giống với cách ghép 2 cột như trên.


### Cách 1: Sử dụng “toán tử &”


Tương tự như cách làm trên, chúng ta sẽ tiến hành theo công thức:=Chuỗi1&Chuỗi 2&….


Chúng ta có thể để các chuỗi nối dính liền vào nhau thì giữa 2 dấu (“”) sẽ viết liền hoặc giữa các chuỗi có khảng cách trắng thì giữa 2 dấu (” “) sẽ để khoảng trắng. Sau khi thực hiện nhập thông tin hoàn chỉnh sẽ tạo thành ô D2. Ô này được gộp từ 3 ô A2, B2, C2.


![](https://hddtvn.com/wp-content/uploads/2021/01/opMw8dk.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


Để thực hiện với các ô còn lại một cách nhanh chóng, chúng ta chỉ cần copy công thức vào các ô tiếp theo để có được cột Dữ liệu hoàn chỉnh.


![](https://hddtvn.com/wp-content/uploads/2021/01/c1wjgdD.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


### Cách 2: Sử dụng hàm CONCATENATE


Chúng ta tiến hành nối chuỗi CONCATENATE trong Excel theo công thức **=CONCATENATE(text1,text2,…)**.


Nếu người dùng để các chuỗi nối dính liền vào nhau thì giữa 2 dấu (“”) sẽ viết liền. Nếu để giữa các chuỗi có khoảng cách trắng thì giữa 2 dấu (” “) sẽ để khoảng trắng.


![](https://hddtvn.com/wp-content/uploads/2021/01/yxwvub2.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


Cuối cùng, bạn chỉ cần kéo từ ô đầu tiên xuống những ô còn lại. Ta được kết quả như sau:


![Hướng dẫn cách ghép 2 hay nhiều cột trong Excel](https://hddtvn.com/wp-content/uploads/2021/01/dNxyhJh.png "Hướng dẫn cách ghép 2 hay nhiều cột trong Excel")


Qua bài viết trên, bạn có thể thực hiện việc ghép 2 hay nhiều cột thành 1 cột duy nhất. Mong rằng bài viết này sẽ đem lại thông tin hữu ích cho bạn đọc. Chúc bạn thành công!


moreGộp 2 (hay nhiều cột) thành 1 cột duy nhất là thao tác thường xuyên…

