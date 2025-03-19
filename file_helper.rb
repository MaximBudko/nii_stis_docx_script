module FileHelper
  def self.safe_copy(source, target, max_retries = 5)
    retries = 0
    begin
      # Force close any open handles
      GC.start
      sleep(0.2)
      
      # Try to remove existing file
      if File.exist?(target)
        File.chmod(0777, target) rescue nil
        File.delete(target) rescue nil
      end
      
      # Copy file
      FileUtils.copy_file(source, target)
      File.chmod(0666, target) rescue nil
      true
    rescue => e
      retries += 1
      if retries < max_retries
        sleep(1)
        retry
      end
      raise e
    end
  end

  def self.safe_move(source, target, max_retries = 5)
    retries = 0
    begin
      GC.start
      sleep(0.2)

      if File.exist?(target)
        File.chmod(0777, target) rescue nil
        File.delete(target) rescue nil
      end

      if RUBY_PLATFORM =~ /mswin|mingw|windows/
        # Windows-specific move
        File.rename(source, target)
      else
        # Unix-style move
        FileUtils.mv(source, target)
      end
      true
    rescue => e
      retries += 1
      if retries < max_retries
        sleep(1)
        retry
      end
      raise e
    end
  end
end
