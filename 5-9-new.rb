# frozen_string_literal: true

require 'faraday'
require 'json'
require 'zip'
require 'htmltoword'
require 'google_drive'

class UserAPIRESTful
  attr_accessor :name, :sex, :active, :created_at, :avatar

  def initialize(url)
    @url = url
    @connection = Faraday.new(url: @url)
  end

  def set_user_info(name, sex, active, created_at, avatar)
    @name = name
    @sex = sex
    @active = active
    @created_at = created_at
    @avatar = avatar
  end

  def delete_user(user_id)
    response = @connection.delete("/users/#{user_id}")

    if response.status == 200
      puts "User with ID #{user_id} is deleted."
    elsif response.status == 404
      puts "not found user with ID #{user_id}."
    end
  end

  def update_user(user_id)
    user_data = {
      name: @name,
      sex: @sex,
      active: @active,
      created_at: @created_at,
      avatar: @avatar
    }

    response = @connection.put("/users/#{user_id}") do |req|
      req.headers['Content-Type'] = 'application/json'
      req.body = user_data.to_json
    end

    if response.status == 200
      user = JSON.parse(response.body)
      puts 'user information is updated:'
      puts user
    elsif response.status == 404
      puts "not found user with ID #{user_id}."
    end
  end

  def create_user
    user_data = {
      name: @name,
      sex: @sex,
      active: @active,
      created_at: @created_at,
      avatar: @avatar
    }

    response = @connection.post do |req|
      req.headers['Content-Type'] = 'application/json'
      req.body = user_data.to_json
    end

    if response.status == 201
      user = JSON.parse(response.body)
      puts user
    elsif response.status == 400
      puts 'error'
    end
  end

  def get_active_users
    response = @connection.get('/users?active=true')

    if response.status == 200
      users = JSON.parse(response.body)
      puts 'List user active:'
      users.each do |user|
        puts user
      end
    else
      puts 'error'
    end
  end

  def export_users_to_file(file_path)
    response = @connection.get('/users')

    if response.status == 200
      users = JSON.parse(response.body)
      thead_html = <<-HTML
        <thead>
          <tr>
            <th>Name</th>
            <th>Sex</th>
            <th>active</th>
            <th>created_at</th>
            <th>Avatar</th>
          </tr>
        </thead>
      HTML
      tbody_html = ''
      users.each do |user|
        tbody_html += <<-HTML
          <tr>
            <td>#{user['name']}</td>
            <td>#{user['sex']}</td>
            <td>#{user['active']}</td>
            <td>#{user['created_at']}</td>
            <td>#{user['avatar']}</td>
          </tr>
        HTML
      end
      table_html = "<table>#{thead_html}<tbody>#{tbody_html}</tbody></table>"

      word_file_path = "#{file_path}.docx"
      Htmltoword::Document.create_and_save(table_html, word_file_path)

      @zip_file_path = "#{file_path}.zip"
      Zip::File.open(@zip_file_path, Zip::File::CREATE) do |zipfile|
        zipfile.add('user_info.docx', word_file_path)
      end

      File.delete(word_file_path)

      puts "Success #{@zip_file_path}"
    else
      puts 'Error'
    end
  end

  def upload_to_drive
    conn = Faraday.new(url: 'https://accounts.google.com')

    response = conn.get(
      '/o/oauth2/v2/auth', {
        scope: 'https://www.googleapis.com/auth/drive',
        access_type: 'offline',
        include_granted_scopes: 'true',
        response_type: 'code',
        state: 'state_parameter_passthrough_value',
        redirect_uri: 'http://localhost:3002/',
        client_id: '803975415585-bu9e1d9idupoo1is8utqnuhl90lk27g9.apps.googleusercontent.com'
      }
    )
    response.env.url

    # xử lý callback ở đây sẽ có code để tạo refresh token

    code = '4/0Adeu5BVka89NCIHqbj4tGtakQ8Vizs4yWggVxP-cDBwCEk4ww-HHEMyyfCpJ81_OWTLbBg'

    post_conn = Faraday.new(url: 'https://oauth2.googleapis.com')

    post_response = post_conn.post(
      '/token', {
        client_id: '803975415585-bu9e1d9idupoo1is8utqnuhl90lk27g9.apps.googleusercontent.com',
        client_secret: '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',
        code: code,
        grant_type: 'authorization_code',
        redirect_uri: 'http://localhost:3002/'
      }
    )

    access_token = eval(post_response.body)[:access_token]

    session = GoogleDrive::Session.login_with_oauth(access_token)

    session.upload_from_file(@zip_file_path, File.basename(@zip_file_path), convert: false)
  end
end

url = 'https://6418014ee038c43f38c45529.mockapi.io/api/v1/users'
user_creator = UserAPIRESTful.new(url)

user_creator.get_active_users

user_creator.set_user_info('Pike', 'male', true, Time.now, 'https://cdn.popsww.com/blog/sites/2/2022/02/megumin.jpg')
user_creator.create_user
user_creator.set_user_info('Peki', 'male', true, Time.now, 'https://cdn.popsww.com/blog/sites/2/2022/02/megumin.jpg')
user_creator.update_user(69)
user_creator.delete_user(69)
user_creator.export_users_to_file('user-information')
