package simpleexcel

type ExcelError struct {
	Cause string
	Stack *string
}

func NewExcelError(cause string, e error) *ExcelError {
	stack := new(string)
	if e != nil {
		*stack = e.Error()
	} else {
		stack = nil
	}
	return &ExcelError{
		Cause: cause,
		Stack: stack,
	}
}
func (e ExcelError) Error() string {
	if e.Stack != nil {
		return e.Cause + "\n" + *e.Stack
	} else {
		return e.Cause
	}
}
