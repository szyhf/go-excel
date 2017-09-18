package excel

import (
	"github.com/szyhf/go-excel/internal"
)

type Connecter = internal.Connecter
type Reader = internal.Reader

func NewConnecter() Connecter {
	return internal.NewConnect()
}
