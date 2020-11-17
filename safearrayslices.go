// +build windows

package ole

import (
	"fmt"
	"unsafe"
)

func safeArrayFromByteSlice(slice []byte) *SafeArray {
	array, _ := safeArrayCreateVector(VT_UI1, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []byte to SAFEARRAY")
	}

	for i, v := range slice {
		err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(&v)))
		if err != nil {
			fmt.Println("safeArrayFromByteSlice", err)
		}
	}
	return array
}

func SafeArrayFromStringSlice(slice []string) *SafeArray {
	array, _ := safeArrayCreateVector(VT_BSTR, 0, uint32(len(slice)))

	if array == nil {
		panic("Could not convert []string to SAFEARRAY")
	}
	// SysAllocStringLen(s)
	for i, v := range slice {
		err := safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(SysAllocStringLen(v))))
		if err != nil {
			fmt.Println("safeArrayFromStringSlice", err)
		}
	}
	return array
}

func SafeArrayFromVariantSlice(slice []interface{}) *SafeArray {
	array, err := safeArrayCreateVector(VT_INT_PTR, 0, uint32(len(slice)))

	if array == nil {
		fmt.Print(err)
		panic("Could not convert []*VARIANT to SAFEARRAY")
	}
	// SysAllocStringLen(s)
	for i, v := range slice {
		fmt.Printf("SafeArrayFromVariantSlice %p\n", v.(*IDispatch))
		fmt.Printf("SafeArrayFromVariantSlice %#x\n", uintptr(unsafe.Pointer((v.(*IDispatch)))))

		err = safeArrayPutElement(array, int64(i), uintptr(unsafe.Pointer(v.(*IDispatch))))
		if err != nil {
			fmt.Println(err)
		}
	}
	return array
}
