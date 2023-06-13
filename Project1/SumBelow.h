#pragma once

namespace Project1 {

	using namespace System::Windows::Forms;

	public ref struct SumBelow {
		DataGridViewCell^ cell;
		int belowCellsCount;

		SumBelow(DataGridViewCell^ cell, int belowCellsCount) {
			this->cell = cell;
			this->belowCellsCount = belowCellsCount;
		}
	};
}