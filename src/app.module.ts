import { Module } from '@nestjs/common';
import { ExcelController } from './excel/excel.controller';
import { ExcelService } from './excel/excel.service';

@Module({
  imports: [],
  controllers: [ExcelController],
  providers: [ExcelService],
})
export class AppModule {}
