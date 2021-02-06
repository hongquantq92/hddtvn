Hướng dẫn sửa các lỗi thường gặp trên phần mềm HTKK
===================================================

[:gift:Xem chi tiết ở đây:gift:](https://hddtvn.com/huong-dan-sua-cac-loi-thuong-gap-tren-phan-mem-htkk/)Đối với kế toán, phần mềm HTKK là một công cụ hỗ trợ công việc không thể thiếu. Trong quá trình cài đặt và sử dụng, người dùng sẽ khó tránh khỏi một số lỗi như: Gõ không được tiếng việt, không đăng nhập được, lỗi HTKK không nhận hết dữ liệu từ file excel,.. …
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Đối với kế toán, phần mềm HTKK là một công cụ hỗ trợ công việc không thể thiếu. Trong quá trình cài đặt và sử dụng, người dùng sẽ khó tránh khỏi một số lỗi như: Gõ không được tiếng việt, không đăng nhập được, lỗi HTKK không nhận hết dữ liệu từ file excel,.. Đây chỉ là những lỗi phần mềm cơ bản. Bài viết dưới đây sẽ hướng dẫn bạn một vài kỹ thuật để khắc phục lỗi này.


1. Lỗi: Run Time Error z‘9′: Subscript out of range
---------------------------------------------------


– Nguyên nhân: Lỗi này gây ra bởi bạn để sai định dạng ngày tháng.


– Cách sửa:


+ Cài đặt lại định dạng ngày cho hệ thống.


+ Bằng cách kích chuột vào Start => Control Panel->Region and language->chọn Tab “Fomats”->Bấm “Additional Settings” ->chọn Tab “Date” chỉnh lại Short date: dd/MM/yyyy; Long date: dddd, dd MMMM, yyyy.


+ Chuyển qua Tab “Location” chọn Home location: United States => OK để hoàn thành.


![HTKK lỗi run-time error](https://hddtvn.com/wp-content/uploads/2021/01/HTKK-7.jpg)


2. Lỗi: Run-time error ’75′ : Path/file access error
----------------------------------------------------


– Nguyên nhân:


+ Do tài khoản người dùng (User) ứng với hệ điều hành trên máy tính của bạn khi Log on vào Windows không có quyền ghi (Write) lên bất kỳ thư mục do tài khoản Administrator tạo ra vì vậy khi bạn khởi tạo chương trình sẽ bị lỗi như trên.


+ Hoặc cũng có thể bạn cài đặt HTKK vào thư mục có tên chứa dấu tiếng việt.


– Cách sửa:


+ Truy cập vào Window bằng tài khoản có toàn quyền (Administrator) và phân lại quyền cho thư mục HTKK có quyền ghi (write). Cụ thể thực hiện như sau:




* Bước 1: Tìm đến thư mục có tên HTKK sau đó nhấn nút phải chuột ở thư mục HTKK và chọnmenu properties. Thường thư mục này được tìm thấy theo đường dẫn C:\Program Files (x86)\HTKK hoặc C:\Program Files\HTKK.

* Bước 2: Ta chọn thẻ Security ở cửa sổ Properties, ta thục hiện như sau:



                     -> Tại mục Group or User names bạn tích chuột vào dòng Administrator.


                     -> Tại mục Permissions for Administrator ta chọn dấu tích vào phần ô vuông của “Full Control” ở cột Allow.




* Bước 3: Kích chuột vào nút OK để hoàn thành.



+ Trường hợp Bạn cài đặt HTKK trong thư mục có tên tiếng việt có dấu các bạn thực hiện bỏ dấu tiếng Việt và đảm bảo tên thư mục đó không chứa các ký tự đặc biệt như: ! @ # $ % …


3. Lỗi xảy ra khi kích chuột vào nút Ghi có thông báo kê khai sai
-----------------------------------------------------------------


– Nguyên nhân: Lỗi này xảy ra nhiều khả năng là do bạn tích chọn nhiều hoặc tất cả các phụ lục mà lại chỉ thực hiện kê khai trên bảng kê Bán ra, mua vào nên khi lưu Phần mềm thông báo lỗi.


– Cách sửa:


+ Khi thực hiện kê khai, bạn chỉ tích vào những phụ lục cần thiết áp dụng cho doanh nghiệp mình.


+ Trường hợp Bạn đã tích chọn tất cả các lựa chọn thì bạn thực hiện xóa lần lượt từng phụ lục thừa thì lỗi này sẽ được khắc phục.


4. Lỗi không gõ được tiếng Việt trong HTKK
------------------------------------------


![phần mềm HTKK lỗi tiếng TV](https://hddtvn.com/wp-content/uploads/2021/01/htkk-3-4-3.jpg)


– Nguyên nhân: Lỗi này có thể do bạn chưa thiết lập đúng định dạng gõ chuẩn với HTKK trong phần mềm gõ tiếng Việt. Lỗi này khá cơ bản và phổ biến nhất.


– Cách sửa:


+ Với phần mềm Unikey: Trong cửa sổ làm việc của phần mềm Unikey người sử dụng chọn nút “Mở rộng” sau đó tích chọn vào chức năng “Sử dụng clipboard cho unicode” rồi bấm “Đóng” để áp dụng.


+ Với phần mềm Vietkey: Trong cửa sổ làm việc của phần mềm Vietkey người sử dụng chọn tab Bảng mã sau đó tích chọn vào chức năng “Unicode dựng sẵn”.


– Chú ý:


Sau khi xử lý xong công việc, Bạn cần đưa các thiết lập trở về thời điểm ban đầu đối với các phần mềm hỗ trợ gõ tiếng Việt trên, để tránh xảy ra lỗi trong quá trình soạn thảo văn bản với các phần mềm khác.


5. Lỗi không gõ được tiếng Việt sau khi cài phiên bản HTKK mới nhất
-------------------------------------------------------------------


Sau khi cài xong phiên bản HTKK mới nhất, bạn thường không gõ được tiếng Việt, để xử lý vấn đề này bạn lần lượt làm theo hướng dẫn theo thứ tự từ nhỏ đến lớn của các ô hình chữ nhật mầu đỏ ở các hình minh họa bên dưới là ổn. Chú ý, hướng dẫn này chỉ áp dụng với những máy tính đang sử dụng phần mềm gõ tiếng Việt Unikey.


– Đầu tiên các bạn click chuột phải vào biểu tượng Unikey trên màn hình Window


– Chọn “Properties” => Chọn “Compatibility” => Tích chọn vào ô ” Run this program as an administrator” => nhấn nút ” OK”.


6. Lỗi không đăng nhập được vào phần mềm HTKK
---------------------------------------------


– Lỗi này xảy ra khi bạn nhập Mã số thuế doanh nghiệp vào HTKK và nhấn nút Đồng Ý, nhưng nó không truy cập vào hệ thống.


– Cách sửa: Vào Menu Start -> Run -> gõ lệnh: regsvr32 scrrun.dll;


7. Lỗi: Datafile bị lỗi lệch chỉ tiêu với tờ khai
-------------------------------------------------


– Nguyên nhân:


Do dữ liệu có trong máy tính của bạn thuộc phiên bản cũ, trong khi phiên bản HTKK mới đã sửa đổi mẫu tờ khai mới nên sảy ra lỗi trên. Lỗi này thường gặp khi nâng cấp các phiên bản HTKK mới có sửa đổi theo mẫu của thông tư mới.


– Cách sửa:


+ Khi bản mới có sự khác biệt về tờ khai, thường sẽ cho phép cài đồng thời cả hai phiên bản HTKK (cũ và mới) đề theo dõi số liệu cũ và làm mới. Vì vậy Bạn không cần gỡ bản cũ mà chỉ cài thêm bản mới để làm.


+ Trong trường hợp này tuyệt đối bạn không được lấy dữ liệu cũ đổ vào phiên bản HTKK mới để tránh gây ra lỗi trên.


8. Lỗi HTKK không nhận hết dữ liệu từ bảng kê excel
---------------------------------------------------


– Nguyên nhân:


Lỗi này xảy ra khi nhận dữ liệu bảng kê hoá đơn mua vào, bán ra từ phần mềm kế toán, đã copy dữ liệu vào mẫu định dạng chuẩn lấy từ phần “Trợ giúp” của ứng dụng HTKK, dữ liệu không nhận vào ứng dụng HTKK hoặc nhận không đúng các chỉ tiêu trên bảng kê.


– Cách sửa:


Bạn cần làm lại dữ liệu bảng tính Excel đó và không được làm thay đổi cấu trúc các cột của bảng phải để nguyên theo mẫu của chương trình, Bạn chỉ có thể thêm dòng hoặc xoá dòng và phải điền đầy đủ thông tin, kể cả cột số thứ tự.


![lỗi không nhận dữ liệu](https://hddtvn.com/wp-content/uploads/2021/01/loi-bao-cao-thue-htkk.png)


9. Microsoft .NET Framework 3.5 needs to be installed for this installatation to continue
-----------------------------------------------------------------------------------------


– Nguyên nhân:


+ Lỗi này xảy ra khi bạn tiến hành cài đặt HTKK trên hệ điều hành Window, đặc biệt là Win10 (thỉnh thoảng xảy ra ở phiên bản Win8).


+ Lý do xảy ra lỗi này là do hệ điều hành của bạn chưa được cài ứng dụng Microsoft .NET Framework 3.5 – một phần mềm cần thiết để hỗ trợ cho HTKK chạy dược.


– Cách sửa lỗi:


+ Bước 1: Tải phần mềm Microsoft .NET Framework 3.5 từ link sau: https://www.microsoft.com/en-us/download/details.aspx?id=21


+ Bước 2: Tiến hành cài đặt Microsoft .NET Framework 3.5 bằng cách kích chuột vào tệp tin vừa tải được về ở bước 1 và làm theo các chỉ thị trong đó.


+ Bước 3: Tiến hành cài đặt HTKK phiên bản mới nhất theo hướng dẫn ở Phần 1 bên trên cùng.


10. Lỗi “đóng file excel template đang mở trước khi import vào chương trình”
----------------------------------------------------------------------------


– Nguyên nhân:


+ Chưa đóng file excel bảng kê.


+ Sử dụng file Excel mẫu bảng kê không đúng định dạng.


+ Đặt tên file, tên thư mục chứa file có dấu tiếng Việt.


– Cách sửa:


+ Đóng file bảng kê Excel trước khi import dữ liệu từ file Excel này vào phần mềm HTKK. Sau đó, bạn cần sử dụng file Excel mẫu trong phần mềm HTKK (Vào Menu Trợ giúp -> Mẫu bảng kê để lấy).


+ Bạn lưu ý không được đặt tên file, tên thư mục chứa file bằng Tiếng việt có dấu.


11. Lỗi: Run – time error “380″ Invalid property value
------------------------------------------------------


– Nguyên nhân:


Do bạn chưa thiết lập chế độ ngày tháng của máy tính theo dạng dd/mm/yyyy mà HTKK sử dụng.


– Cách sửa:


Kích chọn Start => Settings => Control Panel => Date and time chỉnh đúng theo định dạng dd/mm/yyyy => Chọn OK để hoàn thành.


Như vậy, chúng tôi vừa hỗ trợ bạn sửa các lỗi thường gặp trên phần mềm HTKK. Nếu bạn gặp phải những lỗi nằm ngoài hướng dẫn trên và chưa có cách khắc phục thì hãy để lại bình luận phía dưới để được hỗ trợ nhé. Chúc bạn thành công!



moreĐối với kế toán, phần mềm HTKK là một công cụ hỗ trợ công việc…

